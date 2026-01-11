# 3_Receptura_fixed.py
# - Wczytywanie receptury z arkusza (selectbox + przycisk "Wczytaj")
# - Domieszki traktowane jak normalne składniki (udział objętościowy, bilans, koszt, CO2)
# - Odporne parsowanie liczb z przecinkiem i odstępami (również w data_editor)
# - Stabilne ID w edytorach (brak utraty fokusa)
# - Kruszywo: USUNIĘTA autokorekta / „wyrównywanie” inputów frakcji (normalize tylko do obliczeń)
# - Zapis do Google Sheets + __SUMMARY__

import math
from typing import List, Dict, Any, Optional
from datetime import datetime

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ---------------- Google Sheets auth ----------------
GS_READY = False
try:
    from google.oauth2.service_account import Credentials
    import gspread
    try:
        from gspread_dataframe import get_as_dataframe
    except Exception:
        get_as_dataframe = None

    GSA = "gcp_service_account"
    SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")
    SHEET_MATERIALS = st.secrets.get("SHEET_MATERIALS", "materials")
    SHEET_RECIPES = st.secrets.get("SHEET_RECIPES", "receptury")

    if GSA in st.secrets and SPREADSHEET_ID:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        CREDS = Credentials.from_service_account_info(st.secrets[GSA], scopes=SCOPES)
        GS_READY = True
except Exception:
    GS_READY = False
    get_as_dataframe = None


# ---------------- Helpers: state init ----------------
def init_state(key: str, default):
    if key not in st.session_state:
        st.session_state[key] = default


# ---------------- Helpers: liczby z przecinkiem / odstępami ----------------
def to_num_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
        .str.replace("\xa0", " ", regex=False)  # NBSP
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    )


def to_num_val(x) -> float:
    """Odporne parsowanie pojedynczej wartości (np. z data_editor), obsługa przecinka i spacji."""
    if x is None:
        return math.nan
    if isinstance(x, (int, float)) and not (isinstance(x, float) and math.isnan(x)):
        return float(x)
    s = (
        str(x)
        .replace("\xa0", " ")  # NBSP
        .replace(" ", "")
        .replace(",", ".")
        .strip()
    )
    return pd.to_numeric(s, errors="coerce")


# ---------------- Streamlit Page Setup ----------------
st.set_page_config(page_title="Receptura", page_icon="", layout="wide")
st.title("Receptura mieszanki betonowej")

# ---------------- Stałe i kolejność kategorii ----------------
CATEGORIES_ORDERED: List[tuple[str, str]] = [
    ("spoiwo", "Spoiwo"),
    ("dodatek", "Dodatki mineralne"),
    ("kruszywo", "Kruszywo"),
    ("woda", "Woda"),
    ("domieszka", "Domieszki chemiczne"),
]

# ---------------- Nagłówki eksportu receptury ----------------
HEADERS = [
    "timestamp", "recipe_name",
    "material_id", "nazwa", "kategoria",
    "gestosc_kgm3", "udzial_pct", "obj_m3", "masa_kgm3",
    "sum_obj_m3m3", "sum_mas_kgm3", "gestosc_mix_kgm3", "w_c",
    "fck_mpa", "fctm_mpa", "ecm_gpa"
]


# ----------------- Funkcje pomocnicze (tabele) -----------------
def ensure_table_columns(tbl: pd.DataFrame) -> pd.DataFrame:
    t = tbl.copy()
    if "id" not in t.columns:
        t["id"] = pd.NA
    if "nazwa" not in t.columns:
        t["nazwa"] = ""
    if "kategoria" not in t.columns:
        t["kategoria"] = pd.NA

    if "gestosc_kgm3" not in t.columns:
        if "gestosc_gcm3" in t.columns:
            t["gestosc_kgm3"] = to_num_series(t["gestosc_gcm3"]) * 1000.0
        else:
            t["gestosc_kgm3"] = pd.NA

    if "udzial_pct" not in t.columns:
        for alt in ["udział_%", "Udział [%]", "udzial_proc", "udział %", "udzial %", "Udział objętościowy [%]"]:
            if alt in t.columns:
                t["udzial_pct"] = to_num_series(t[alt]).fillna(0.0)
                break
        else:
            t["udzial_pct"] = 0.0

    t["id"] = pd.to_numeric(t["id"], errors="coerce")
    t["udzial_pct"] = pd.to_numeric(t["udzial_pct"], errors="coerce").fillna(0.0)
    t["gestosc_kgm3"] = pd.to_numeric(t["gestosc_kgm3"], errors="coerce")

    tech_cols = ["id", "nazwa", "kategoria", "gestosc_kgm3", "udzial_pct"]
    rest = [c for c in t.columns if c not in tech_cols]
    return t[tech_cols + rest]


def apply_edited_back(original_tbl: pd.DataFrame, edited_tbl: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
    orig = original_tbl.copy()
    ed = edited_tbl.copy()
    if "id" not in ed.columns and "id" in orig.columns:
        ed["id"] = orig["id"].values

    orig_idx = orig.set_index("id", drop=False)
    ed_idx = ed.set_index("id", drop=False)

    orig_idx = orig_idx[~orig_idx.index.duplicated(keep="first")]
    ed_idx = ed_idx[~ed_idx.index.duplicated(keep="first")]

    num_cols = set(orig.select_dtypes(include=["number"]).columns.tolist())

    for inp_col, tech_col in mapping.items():
        if inp_col not in ed_idx.columns:
            continue
        common = orig_idx.index.intersection(ed_idx.index)
        if common.empty:
            continue
        for i in common:
            new_val = ed_idx.at[i, inp_col]
            if pd.isna(new_val) or (isinstance(new_val, str) and new_val.strip() == ""):
                continue
            if tech_col in num_cols:
                coerced = pd.to_numeric(new_val, errors="coerce")
                if pd.notna(coerced):
                    orig_idx.at[i, tech_col] = coerced
            else:
                orig_idx.at[i, tech_col] = new_val

    return orig_idx.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def read_materials_from_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    if not GS_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets w secrets.")
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    try:
        ws = ss.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        raise RuntimeError(f"Nie znaleziono zakładki \"{sheet_name}\" w tym pliku.")

    if get_as_dataframe is not None:
        df = get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
    else:
        rows = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(rows)

    for c in ["id", "nazwa", "kategoria", "gestosc_gcm3", "cena_pln", "co2e_kgkg"]:
        if c not in df.columns:
            df[c] = None
    if "cena_za" not in df.columns:
        df["cena_za"] = "kg"

    df["id"] = to_num_series(df["id"])
    df["gestosc_gcm3"] = to_num_series(df["gestosc_gcm3"])
    df["cena_pln"] = to_num_series(df["cena_pln"])
    df["co2e_kgkg"] = to_num_series(df["co2e_kgkg"])

    df["cena_za"] = df["cena_za"].astype(str).str.strip().str.lower()
    df["kategoria"] = df["kategoria"].astype(object)
    df["gestosc_kgm3"] = df["gestosc_gcm3"] * 1000.0

    def _price_per_kg(row):
        price = row["cena_pln"]
        unit = str(row.get("cena_za", "")).strip().lower()
        rho_gcm3 = row["gestosc_gcm3"]
        if pd.isna(price):
            return math.nan

        unit_norm = unit.replace("/", "").replace(" ", "")
        if unit_norm in ("kg", "kilogram"):
            return price

        if unit_norm in ("l", "litr", "litry", "dm3"):
            if pd.isna(rho_gcm3) or rho_gcm3 <= 0:
                return math.nan
            return price / rho_gcm3

        return price

    df["cena_pln"] = df.apply(_price_per_kg, axis=1)
    df["nazwa"] = df["nazwa"].astype("object").where(pd.notna(df["nazwa"]), "")

    return df


def options_for_category(df: pd.DataFrame, cat: str) -> Dict[int, str]:
    sdf = df.loc[df["kategoria"] == cat, ["id", "nazwa"]].dropna(subset=["id"]).copy()
    if sdf.empty:
        return {}
    sdf["id"] = pd.to_numeric(sdf["id"], errors="coerce").astype("Int64").dropna().astype(int)
    sdf = sdf.drop_duplicates(subset=["id"], keep="first").sort_values("id")
    return dict(zip(sdf["id"].tolist(), sdf["nazwa"].astype(str).tolist()))


def build_table_for_selection(df_all: pd.DataFrame, ids: List[int], cat_key: str) -> pd.DataFrame:
    ids = [int(x) for x in ids]
    if not ids:
        return pd.DataFrame(columns=["id", "nazwa", "kategoria", "gestosc_kgm3", "udzial_pct"])
    t = df_all[df_all["id"].isin(ids)].copy()
    t = ensure_table_columns(t)
    t["udzial_pct"] = pd.to_numeric(t.get("udzial_pct", 0.0), errors="coerce").fillna(0.0)
    t = t.sort_values("id").drop_duplicates(subset=["id"], keep="first").reset_index(drop=True)
    return t


# ============================================================
# Receptury: wczytanie z arkusza + odtworzenie session_state
# ============================================================
@st.cache_data(show_spinner=False)
def read_recipes_rows(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    if not GS_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets w secrets.")
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    ws = ss.worksheet(sheet_name)
    values = ws.get_all_values()

    if not values:
        return pd.DataFrame(columns=HEADERS)

    header = values[0]
    df = pd.DataFrame(values[1:], columns=header)

    for c in HEADERS:
        if c not in df.columns:
            df[c] = ""

    df["recipe_name"] = df["recipe_name"].astype(str)
    df["nazwa"] = df["nazwa"].astype(str)
    df["kategoria"] = df["kategoria"].astype(str)

    for c in [
        "material_id", "gestosc_kgm3", "udzial_pct", "obj_m3", "masa_kgm3",
        "sum_obj_m3m3", "sum_mas_kgm3", "gestosc_mix_kgm3", "w_c",
        "fck_mpa", "fctm_mpa", "ecm_gpa"
    ]:
        if c in df.columns:
            df[c] = to_num_series(df[c])

    return df


def apply_loaded_recipe(df_loaded: pd.DataFrame, df_all: pd.DataFrame):
    for cat_key, _ in CATEGORIES_ORDERED:
        st.session_state[f"msel_{cat_key}"] = []
        st.session_state.pop(f"tbl_{cat_key}", None)

    existing_ids = set(df_all["id"].astype(int).tolist())

    for cat_key, _ in CATEGORIES_ORDERED:
        sdf = df_loaded[df_loaded["kategoria"].astype(str).str.strip().eq(cat_key)].copy()
        if sdf.empty:
            continue

        sdf["id"] = pd.to_numeric(sdf["material_id"], errors="coerce")
        sdf = sdf.dropna(subset=["id"]).copy()
        sdf["id"] = sdf["id"].astype(int)
        sdf = sdf[sdf["id"].isin(existing_ids)].copy()
        if sdf.empty:
            continue

        st.session_state[f"msel_{cat_key}"] = sdf["id"].tolist()

        base_tbl = build_table_for_selection(df_all, st.session_state[f"msel_{cat_key}"], cat_key)
        base_tbl = ensure_table_columns(base_tbl)

        ud_map = dict(zip(
            sdf["id"].astype(int).tolist(),
            pd.to_numeric(sdf["udzial_pct"], errors="coerce").fillna(0.0).tolist()
        ))
        base_tbl["udzial_pct"] = base_tbl["id"].astype(int).map(ud_map).fillna(0.0)

        if cat_key == "kruszywo":
            total_pct = float(pd.to_numeric(base_tbl["udzial_pct"], errors="coerce").fillna(0.0).sum())
            st.session_state["kruszywo_total_pct"] = round(total_pct, 4)
            if total_pct > 0:
                base_tbl["share_in_agg_pct"] = (base_tbl["udzial_pct"] / total_pct) * 100.0
            else:
                n = max(len(base_tbl), 1)
                base_tbl["share_in_agg_pct"] = 100.0 / n

        st.session_state[f"tbl_{cat_key}"] = base_tbl.reset_index(drop=True).copy()


# ---------------- Przyciski / odświeżenie cache ----------------
refresh_col, _ = st.columns([1, 3])
with refresh_col:
    if st.button("🔄 Odśwież dane"):
        st.cache_data.clear()
        st.toast("Dane odświeżone.")


# ---------------- Wczytanie materiałów ----------------
if not GS_READY:
    st.error(
        "Brak konfiguracji Google Sheets. Uzupełnij w `.streamlit/secrets.toml`:\n"
        "SPREADSHEET_ID, SHEET_MATERIALS oraz blok [gcp_service_account]."
    )
    st.stop()

try:
    df_all = read_materials_from_sheet(SPREADSHEET_ID, SHEET_MATERIALS).copy()
except Exception as e:
    st.error(f"Nie udało się wczytać materiałów z Google Sheets: {e}")
    st.stop()

if df_all.empty:
    st.warning("Arkusz materiałów jest pusty.")
    st.stop()

df_all["id"] = pd.to_numeric(df_all["id"], errors="coerce").astype("Int64")
df_all = df_all.dropna(subset=["id"]).copy()
df_all["id"] = df_all["id"].astype(int)
df_all = df_all.sort_values(["id"]).drop_duplicates(subset=["id"], keep="first").reset_index(drop=True)


# ============================================================
# UI: Wczytaj recepturę z bazy (selectbox + przycisk)
# ============================================================
st.markdown("---")
st.subheader("Wczytaj recepturę z bazy")

try:
    df_recipes_all = read_recipes_rows(SPREADSHEET_ID, SHEET_RECIPES)
except Exception as e:
    df_recipes_all = pd.DataFrame()
    st.error(f"Nie udało się wczytać arkusza receptur: {e}")

if df_recipes_all.empty:
    st.info("Brak zapisanych receptur w arkuszu.")
else:
    recipe_names = sorted(df_recipes_all["recipe_name"].dropna().astype(str).unique().tolist())

    load_col1, load_col2 = st.columns([3, 1], gap="small")

    init_state("sel_recipe_to_load", recipe_names[0] if recipe_names else "")

    with load_col1:
        st.markdown("&nbsp;", unsafe_allow_html=True)
        st.selectbox(
            label="Wybierz recepturę do wczytania",
            options=recipe_names,
            key="sel_recipe_to_load",
            label_visibility="collapsed",
        )

    with load_col2:
        # spacer o wysokości etykiety po lewej, żeby przycisk zrównał się z polem wyboru
        st.markdown("&nbsp;")
        do_load = st.button(
            "⬇️ Wczytaj",
            use_container_width=True,
            key="btn_load_recipe_v2",
        )

    sel_load = st.session_state.get("sel_recipe_to_load")
    if do_load and sel_load:
        df_r = df_recipes_all[df_recipes_all["recipe_name"].astype(str) == str(sel_load)].copy()

        df_sum = df_r[df_r["nazwa"].astype(str).str.upper().isin(["SUMMARY", "__SUMMARY__"])].copy()
        if not df_sum.empty:
            last = df_sum.iloc[-1]
            st.session_state["recipe_name_in"] = str(sel_load)
            if pd.notna(last.get("fck_mpa")):
                st.session_state["fck_in"] = float(last["fck_mpa"])
            if pd.notna(last.get("fctm_mpa")):
                st.session_state["fctm_in"] = float(last["fctm_mpa"])
            if pd.notna(last.get("ecm_gpa")):
                st.session_state["ecm_in"] = float(last["ecm_gpa"])

        df_items = df_r[~df_r["nazwa"].astype(str).str.upper().isin(["SUMMARY", "__SUMMARY__"])].copy()
        apply_loaded_recipe(df_items, df_all)

        st.success(f"Wczytano recepturę: {sel_load}")
        st.rerun()

# ---------------- Nazwa receptury + parametry ----------------
init_state("recipe_name_in", "")
init_state("fck_in", 0.0)
init_state("fctm_in", 0.0)
init_state("ecm_in", 0.0)

with st.container():
    top_left, _ = st.columns([2, 3], gap="small")
    with top_left:
        st.markdown("## Parametry receptury")
        recipe_name = st.text_input("Nazwa receptury", placeholder="np. M1.2", key="recipe_name_in")
        c1, c2, c3 = st.columns([1, 1, 1], gap="small")
        with c1:
            fck_value = st.number_input("fck [MPa]", min_value=0.0, step=0.5, format="%.1f", key="fck_in")
        with c2:
            fctm_value = st.number_input("fctm [MPa]", min_value=0.0, step=0.1, format="%.2f", key="fctm_in")
        with c3:
            ecm_value = st.number_input("Ecm [GPa]", min_value=0.0, step=1.0, format="%.1f", key="ecm_in")


# ---------------- Layout ----------------
col_left, col_right = st.columns([2, 3])

# ============================================================
# LEWA KOLUMNA: wybór + edycja udziałów
# ============================================================
with col_left:
    st.markdown("## Wybór i udziały składników")

    for cat_key, cat_title in CATEGORIES_ORDERED:
        st.subheader(cat_title)

        sel_key = f"msel_{cat_key}"
        init_state(sel_key, [])

        opts = options_for_category(df_all, cat_key)

        current_sel = st.session_state.get(sel_key, [])
        current_sel = [int(x) for x in current_sel if str(x).strip() != "" and str(x).lstrip("-").replace(".", "", 1).isdigit()]
        current_sel = [x for x in current_sel if x in opts.keys()]
        if st.session_state.get(sel_key, []) != current_sel:
            st.session_state[sel_key] = current_sel

        st.multiselect(
            "Wybierz składniki",
            options=list(opts.keys()),
            format_func=lambda x: opts.get(x, str(x)),
            key=sel_key,
            placeholder="Wybierz składniki",
            label_visibility="collapsed",
        )
        sel: List[int] = st.session_state.get(sel_key, [])

        base_tbl = build_table_for_selection(df_all, sel, cat_key)
        skey = f"tbl_{cat_key}"

        base_tbl_i = ensure_table_columns(base_tbl).set_index("id", drop=False)
        base_tbl_i = base_tbl_i[~base_tbl_i.index.duplicated(keep="first")]

        if skey not in st.session_state or not isinstance(st.session_state[skey], pd.DataFrame):
            st.session_state[skey] = base_tbl_i.reset_index(drop=True).copy()
        else:
            cur = ensure_table_columns(st.session_state[skey]).set_index("id", drop=False)
            cur = cur[~cur.index.duplicated(keep="first")]
            cur_u = cur[["udzial_pct"]] if "udzial_pct" in cur.columns else pd.DataFrame(columns=["udzial_pct"])
            merged = base_tbl_i.join(cur_u, how="left", rsuffix="_cur")
            if "udzial_pct_cur" in merged.columns:
                merged["udzial_pct"] = merged["udzial_pct_cur"].combine_first(merged["udzial_pct"])
                merged = merged.drop(columns=["udzial_pct_cur"])
            merged["udzial_pct"] = pd.to_numeric(merged["udzial_pct"], errors="coerce").fillna(0.0)
            st.session_state[skey] = merged.reset_index(drop=True).copy()

        tbl = ensure_table_columns(st.session_state[skey].copy())

        # --- kruszywo special: total + split ---
        if cat_key == "kruszywo":
            if "kruszywo_total_pct" not in st.session_state:
                prev_sum = float(pd.to_numeric(tbl.get("udzial_pct", 0.0), errors="coerce").fillna(0.0).sum())
                st.session_state["kruszywo_total_pct"] = round(prev_sum, 2)

            total_df = pd.DataFrame([{
                "nazwa": "Kruszywo (razem)",
                "Udział objętościowy [%]": float(st.session_state["kruszywo_total_pct"])
            }])

            edited_total = st.data_editor(
                total_df.set_index(pd.Index([int(-1)]), drop=False),
                key="editor_kruszywo_total",
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={"nazwa": st.column_config.TextColumn("nazwa", disabled=True)},
            )
            if edited_total is not None and not edited_total.empty:
                val = to_num_val(edited_total.iloc[0]["Udział objętościowy [%]"])
                st.session_state["kruszywo_total_pct"] = float(val) if pd.notna(val) else 0.0

            # jeśli nie ma splitu, zainicjuj na bazie aktualnych udziałów obj.
            if "share_in_agg_pct" not in tbl.columns:
                prev_udzial = pd.to_numeric(tbl.get("udzial_pct", 0.0), errors="coerce").fillna(0.0)
                total_prev = float(prev_udzial.sum())
                if total_prev > 0:
                    tbl["share_in_agg_pct"] = (prev_udzial / total_prev) * 100.0
                else:
                    n = max(len(tbl), 1)
                    tbl["share_in_agg_pct"] = 100.0 / n
            else:
                tbl["share_in_agg_pct"] = pd.to_numeric(tbl["share_in_agg_pct"], errors="coerce").fillna(0.0)

            split_view = (
                tbl[["id", "nazwa", "share_in_agg_pct"]]
                .rename(columns={"share_in_agg_pct": "Udział w kruszywie [%]"})
                .set_index("id", drop=False)
            )

            edited_split = st.data_editor(
                split_view,
                key="editor_kruszywo_split",
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                    "nazwa": st.column_config.TextColumn("nazwa", disabled=True),
                },
            )

            if edited_split is not None and not edited_split.empty:
                ed = edited_split.rename(columns={"Udział w kruszywie [%]": "share_in_agg_pct"}).copy()
                ed["share_in_agg_pct"] = ed["share_in_agg_pct"].apply(to_num_val).fillna(0.0)
                st.session_state[skey] = apply_edited_back(
                    tbl,
                    ed[["id", "share_in_agg_pct"]],
                    {"share_in_agg_pct": "share_in_agg_pct"}
                )
                tbl = ensure_table_columns(st.session_state[skey].copy())

            # ====== KLUCZOWA ZMIANA: brak autokorekty inputów ======
            tbl["share_in_agg_pct"] = pd.to_numeric(tbl["share_in_agg_pct"], errors="coerce").fillna(0.0)
            share_sum = float(tbl["share_in_agg_pct"].sum())
            total_pct = float(st.session_state.get("kruszywo_total_pct", 0.0))

            # informacja zamiast automatycznej normalizacji w tabeli
            if len(tbl) > 0 and abs(share_sum - 100.0) > 0.5 and share_sum > 0:
                st.info(
                    f"Suma 'Udział w kruszywie' = {share_sum:.2f}%. "
                    "Nie koryguję automatycznie; do obliczeń normalizuję proporcjonalnie."
                )

            # normalizacja WYŁĄCZNIE do wyliczenia udziału objętościowego frakcji
            if share_sum > 0:
                tbl["udzial_pct"] = (tbl["share_in_agg_pct"] / share_sum) * total_pct
            else:
                tbl["udzial_pct"] = 0.0

            st.session_state[skey] = tbl.reset_index(drop=True).copy()
            continue

        # --- wszystkie inne (w tym domieszka): jak normalne ---
        df_for_edit = (
            tbl[["id", "nazwa", "udzial_pct"]]
            .rename(columns={"udzial_pct": "Udział objętościowy [%]"})
            .set_index("id", drop=False)
        )

        edited = st.data_editor(
            df_for_edit,
            key=f"editor_{cat_key}_volume",
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "nazwa": st.column_config.TextColumn("nazwa", disabled=True),
            },
        )

        if edited is not None and not edited.empty:
            ed = edited.rename(columns={"Udział objętościowy [%]": "udzial_pct"}).copy()
            ed["udzial_pct"] = ed["udzial_pct"].apply(to_num_val).fillna(0.0)
            if "id" not in ed.columns:
                ed["id"] = tbl["id"].values
            st.session_state[skey] = apply_edited_back(tbl, ed, {"udzial_pct": "udzial_pct"}).reset_index(drop=True)
        else:
            st.session_state[skey] = tbl.reset_index(drop=True).copy()


# ============================================================
# PRAWA KOLUMNA: bilans, tabele, parametry, zapis, wykresy
# ============================================================
with col_right:
    parts_base = []
    for cat_key, _ in CATEGORIES_ORDERED:
        skey = f"tbl_{cat_key}"
        if skey in st.session_state and isinstance(st.session_state[skey], pd.DataFrame):
            df = ensure_table_columns(st.session_state[skey].copy())
            if df.empty:
                continue
            df["kategoria"] = cat_key
            df["obj_m3"] = pd.to_numeric(df["udzial_pct"], errors="coerce").fillna(0.0) / 100.0
            df["masa_kg_batch"] = df.apply(
                lambda r: (float(r["obj_m3"]) * float(r["gestosc_kgm3"])) if pd.notna(r.get("gestosc_kgm3")) else 0.0,
                axis=1
            )
            parts_base.append(df)

    if not parts_base or all(df.empty for df in parts_base):
        st.warning("Wybierz materiały w lewej kolumnie, aby zobaczyć skład i parametry.")
        st.stop()

    df_recipe_base = pd.concat(parts_base, ignore_index=True, sort=False)
    df_recipe_base["obj_m3"] = pd.to_numeric(df_recipe_base["obj_m3"], errors="coerce").fillna(0.0)
    df_recipe_base["masa_kg_batch"] = pd.to_numeric(df_recipe_base["masa_kg_batch"], errors="coerce").fillna(0.0)

    obj_sum = float(df_recipe_base["obj_m3"].sum())
    sum_pct_base = float(pd.to_numeric(df_recipe_base["udzial_pct"], errors="coerce").fillna(0.0).sum())
    sum_pct_base = sum_pct_base if sum_pct_base > 0 else 1.0

    if obj_sum > 0:
        df_mix = df_recipe_base.copy()
        df_mix["obj_m3_m3"] = df_mix["obj_m3"] / obj_sum
        df_mix["masa_kgm3"] = df_mix["masa_kg_batch"] / obj_sum
    else:
        df_mix = df_recipe_base.copy()
        df_mix["obj_m3_m3"] = 0.0
        df_mix["masa_kgm3"] = 0.0

    sum_obj_per1 = float(df_mix["obj_m3_m3"].sum())
    df_mix["udzial_pct_norm"] = (df_mix["obj_m3_m3"] / sum_obj_per1 * 100.0) if sum_obj_per1 > 0 else 0.0

    woda_m = float(df_mix.loc[df_mix["kategoria"].eq("woda"), "masa_kgm3"].sum())
    spoiwo_m = float(df_mix.loc[df_mix["kategoria"].eq("spoiwo"), "masa_kgm3"].sum())
    dodatki_m = float(df_mix.loc[df_mix["kategoria"].eq("dodatek"), "masa_kgm3"].sum())

    wc = (woda_m / spoiwo_m) if spoiwo_m > 0 else math.nan
    ws_ratio = (woda_m / (spoiwo_m + dodatki_m)) if (spoiwo_m + dodatki_m) > 0 else math.nan

    masa_sum = float(df_mix["masa_kgm3"].sum())
    gestosc_mix = (masa_sum / sum_obj_per1) if sum_obj_per1 > 0 else math.nan

    df_mix["koszt_PLN_m3"] = (
        pd.to_numeric(df_mix["masa_kgm3"], errors="coerce").fillna(0.0)
        * pd.to_numeric(df_mix.get("cena_pln", 0.0), errors="coerce").fillna(0.0)
    )
    df_mix["CO2e_kgm3"] = (
        pd.to_numeric(df_mix["masa_kgm3"], errors="coerce").fillna(0.0)
        * pd.to_numeric(df_mix.get("co2e_kgkg", 0.0), errors="coerce").fillna(0.0)
    )

    total_cost = float(df_mix["koszt_PLN_m3"].sum())
    total_co2 = float(df_mix["CO2e_kgm3"].sum())
    sum_percent = float(pd.to_numeric(df_recipe_base["udzial_pct"], errors="coerce").fillna(0.0).sum())

    def build_table_for_volume(V_out_m3: float, normalize_pct: bool) -> pd.DataFrame:
        factor = (V_out_m3 / obj_sum) if obj_sum > 0 else 0.0

        base = df_recipe_base.copy()
        base["Udział objętościowy [%]"] = (
            (pd.to_numeric(base["udzial_pct"], errors="coerce").fillna(0.0) / (sum_pct_base if normalize_pct else 1.0))
            * (100.0 if normalize_pct else 1.0)
        )
        base["Objętość [m³]"] = base["obj_m3"] * factor
        base["Masa [kg]"] = base["masa_kg_batch"] * factor

        base = base[["id", "nazwa", "kategoria", "Udział objętościowy [%]", "Objętość [m³]", "Masa [kg]"]]

        for c in ["Objętość [m³]", "Masa [kg]"]:
            base[c] = pd.to_numeric(base[c], errors="coerce").round(3)

        mask_pct = pd.to_numeric(base["Udział objętościowy [%]"], errors="coerce").notna()
        base.loc[mask_pct, "Udział objętościowy [%]"] = pd.to_numeric(
            base.loc[mask_pct, "Udział objętościowy [%]"], errors="coerce"
        ).round(3)

        base = base.rename(columns={"id": "ID", "nazwa": "Nazwa", "kategoria": "Kategoria"})
        return base

    st.markdown(f"## Rzeczywisty skład mieszanki — **{obj_sum:.3f} m³**")
    st.dataframe(build_table_for_volume(obj_sum, normalize_pct=False), use_container_width=True, hide_index=True)

    st.markdown("## Skład mieszanki na **1 m³**")
    st.dataframe(build_table_for_volume(1.0, normalize_pct=True), use_container_width=True, hide_index=True)

    st.markdown("## Skład mieszanki na **podaną objętość**")
    init_state("user_liters", 1000.0)
    user_liters = st.number_input("Podana objętość [l]", min_value=0.0, step=10.0, key="user_liters")
    user_m3 = float(user_liters) / 1000.0
    st.dataframe(build_table_for_volume(user_m3, normalize_pct=True), use_container_width=True, hide_index=True)

    st.markdown("## Parametry mieszanki")
    mix_summary = pd.DataFrame([{
        "fck [MPa]": round(float(fck_value), 1) if fck_value is not None else None,
        "fctm [MPa]": round(float(fctm_value), 2) if fctm_value is not None else None,
        "Ecm [GPa]": round(float(ecm_value), 1) if ecm_value is not None else None,
        "w/c [-]": None if math.isnan(wc) else round(wc, 3),
        "w/(c+d) [-]": None if math.isnan(ws_ratio) else round(ws_ratio, 3),
        "Łączna cena [PLN/m³]": round(total_cost, 2),
        "Łączna CO₂e [kg/m³]": round(total_co2, 2),
        "Łączna masa [kg/m³]": round(masa_sum, 1),
        "Suma objętości [m³/m³]": round(sum_obj_per1, 3),
        "Gęstość mieszanki [kg/m³]": None if math.isnan(gestosc_mix) else round(gestosc_mix, 0),
    }])
    st.dataframe(mix_summary, use_container_width=True, hide_index=True)

    if sum_percent <= 0:
        st.warning("Suma udziałów wynosi 0%. Ustaw niezerowe udziały dla wybranych materiałów.")
    elif abs(sum_percent - 100.0) > 0.5:
        st.warning(f"Suma udziałów to **{sum_percent:.2f}%**. Dla pełnego 1 m³ powinna wynosić **100%**.")

    # ============================================================
    # ZAPIS DO GOOGLE SHEETS
    # ============================================================
    def _normalize_name(s: str) -> str:
        return " ".join(str(s).split()).strip().lower()

    def _get_or_create_worksheet(ss: Any, sheet_name: str) -> Any:
        try:
            return ss.worksheet(sheet_name)
        except Exception:
            ws = ss.add_worksheet(title=sheet_name, rows=2000, cols=max(20, len(HEADERS)))
            ws.update("A1", [HEADERS], value_input_option="RAW")
            return ws

    def _sheet_to_df(ws) -> pd.DataFrame:
        vals = ws.get_all_values() or [HEADERS]
        width = len(HEADERS)
        fixed = []
        for row in vals:
            r = list(row)
            if len(r) < width:
                r += [""] * (width - len(r))
            else:
                r = r[:width]
            fixed.append(r)
        if fixed and fixed[0] == HEADERS:
            fixed = fixed[1:]
        return pd.DataFrame(fixed, columns=HEADERS)

    def _update_sheet_atomic(ws, df: pd.DataFrame):
        ws.clear()
        ws.update("A1", [HEADERS], value_input_option="RAW")
        data = df.astype(object).values.tolist()
        if not data:
            return
        CHUNK = 800
        start_row = 2
        start_col = 1
        ncols = len(HEADERS)
        for i in range(0, len(data), CHUNK):
            chunk = data[i:i + CHUNK]
            end_row = start_row + len(chunk) - 1
            end_col = start_col + ncols - 1
            cell_range = (
                f"{gspread.utils.rowcol_to_a1(start_row, start_col)}:"
                f"{gspread.utils.rowcol_to_a1(end_row, end_col)}"
            )
            ws.update(cell_range, chunk, value_input_option="RAW")
            start_row = end_row + 1

    def _rows_for_recipe(
        recipe_name: str,
        df: pd.DataFrame,
        obj_sum_save: float,
        masa_sum_save: float,
        gestosc_mix_save: float,
        wc_val: float,
        fck_val: Optional[float] = None,
        fctm_val: Optional[float] = None,
        ecm_val: Optional[float] = None,
    ) -> List[List[Any]]:
        ts = datetime.now().isoformat(timespec="seconds")
        rows: List[List[Any]] = []
        for _, r in df.iterrows():
            rows.append([
                ts, recipe_name,
                int(r["id"]) if pd.notna(r["id"]) else "",
                str(r.get("nazwa", "")),
                str(r.get("kategoria", "")),
                float(r.get("gestosc_kgm3", 0)) if pd.notna(r.get("gestosc_kgm3", None)) else "",
                float(r.get("udzial_pct", 0)),
                float(r.get("obj_m3", 0)),
                float(r.get("masa_kgm3", 0)),
                "", "", "", "", "", "", ""
            ])
        rows.append([
            ts, recipe_name, "", "__SUMMARY__", "", "", "", "", "",
            float(obj_sum_save),
            float(masa_sum_save),
            float(gestosc_mix_save) if not math.isnan(gestosc_mix_save) else "",
            float(wc_val) if not math.isnan(wc_val) else "",
            float(fck_val) if fck_val is not None else "",
            float(fctm_val) if fctm_val is not None else "",
            float(ecm_val) if ecm_val is not None else "",
        ])
        return rows

    st.markdown("---")
    st.subheader("Zapis receptury")

    disabled_reason = None
    if not recipe_name:
        disabled_reason = "Podaj nazwę receptury."
    elif sum_obj_per1 <= 0:
        disabled_reason = "Suma udziałów musi być > 0%."

    init_state("chk_overwrite_simple", False)

    with st.form(key="save_recipe_form", clear_on_submit=False):
        colL, _ = st.columns([1.2, 2])
        with colL:
            submit_try = st.form_submit_button("💾 Zapisz recepturę", disabled=disabled_reason is not None)

        if disabled_reason:
            st.info(disabled_reason)

        confirm_overwrite = st.checkbox(
            "Nadpisz istniejącą recepturę o tej nazwie, jeśli istnieje",
            value=bool(st.session_state.get("chk_overwrite_simple", False)),
            key="chk_overwrite_simple"
        )

        if submit_try and recipe_name:
            try:
                gc = gspread.authorize(CREDS)
                ss = gc.open_by_key(SPREADSHEET_ID)
                ws_rec = _get_or_create_worksheet(ss, SHEET_RECIPES)

                existing_df = _sheet_to_df(ws_rec)

                rn_norm = _normalize_name(recipe_name)
                exists_mask = existing_df["recipe_name"].astype(str).apply(_normalize_name).eq(rn_norm)
                exists = bool(exists_mask.any())

                df_save = df_mix.copy()
                df_save["obj_m3"] = df_save["obj_m3_m3"]
                df_save["udzial_pct"] = df_save["udzial_pct_norm"]
                df_save = df_save[["id", "nazwa", "kategoria", "gestosc_kgm3", "udzial_pct", "obj_m3", "masa_kgm3"]].copy()

                new_rows = _rows_for_recipe(
                    recipe_name, df_save,
                    obj_sum_save=sum_obj_per1,
                    masa_sum_save=masa_sum,
                    gestosc_mix_save=gestosc_mix,
                    wc_val=wc,
                    fck_val=float(fck_value) if fck_value is not None else None,
                    fctm_val=float(fctm_value) if fctm_value is not None else None,
                    ecm_val=float(ecm_value) if ecm_value is not None else None,
                )
                new_rows_df = pd.DataFrame(new_rows, columns=HEADERS)

                if exists and not confirm_overwrite:
                    st.warning("Receptura o tej nazwie już istnieje. Zaznacz checkbox, aby **nadpisać**, i kliknij ponownie.")
                else:
                    if exists:
                        keep_df = existing_df.loc[~exists_mask].copy()
                        final_df = pd.concat([keep_df, new_rows_df], ignore_index=True)
                        _update_sheet_atomic(ws_rec, final_df)
                        st.success(f"Nadpisano recepturę \"{recipe_name}\" w arkuszu \"{SHEET_RECIPES}\".")
                    else:
                        try:
                            ws_rec.append_rows(new_rows, value_input_option="RAW")
                        except Exception:
                            final_df = pd.concat([existing_df, new_rows_df], ignore_index=True)
                            _update_sheet_atomic(ws_rec, final_df)
                        st.success(f"Zapisano recepturę \"{recipe_name}\" do arkusza \"{SHEET_RECIPES}\".")
            except Exception as e:
                st.error(f"Nie udało się zapisać receptury: {e}")

    # === Wykresy kołowe ===
    st.markdown("## Wykresy składu mieszanki")

    pie_src = df_mix.copy()
    grp = pie_src.groupby("nazwa", dropna=False, as_index=False).agg({
        "obj_m3_m3": "sum",
        "masa_kgm3": "sum"
    })

    vol_sum = float(pd.to_numeric(grp["obj_m3_m3"], errors="coerce").fillna(0.0).sum())
    mass_sum = float(pd.to_numeric(grp["masa_kgm3"], errors="coerce").fillna(0.0).sum())

    grp["Udział objętościowy [%]"] = grp["obj_m3_m3"] / vol_sum * 100.0 if vol_sum > 0 else 0.0
    grp["Udział masowy [%]"] = grp["masa_kgm3"] / mass_sum * 100.0 if mass_sum > 0 else 0.0
    grp = grp.sort_values("Udział objętościowy [%]", ascending=False).reset_index(drop=True)

    col_pie1, col_pie2 = st.columns(2)

    with col_pie1:
        st.markdown("### Udział objętościowy")
        labels = grp["nazwa"].astype(str).tolist()
        values = grp["Udział objętościowy [%]"].astype(float).tolist()
        fig1, ax1 = plt.subplots()
        if sum(values) > 0:
            ax1.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, counterclock=False)
            ax1.axis("equal")
            st.pyplot(fig1, use_container_width=True)
        else:
            st.info("Brak danych do wykresu objętościowego (suma = 0).")

    with col_pie2:
        st.markdown("### Udział masowy")
        labels_m = grp["nazwa"].astype(str).tolist()
        values_m = grp["Udział masowy [%]"].astype(float).tolist()
        fig2, ax2 = plt.subplots()
        if sum(values_m) > 0:
            ax2.pie(values_m, labels=labels_m, autopct="%1.1f%%", startangle=90, counterclock=False)
            ax2.axis("equal")
            st.pyplot(fig2, use_container_width=True)
        else:
            st.info("Brak danych do wykresu masowego (suma = 0).")
