# 3_Receptura_fixed.py
# - Wczytywanie receptury z arkusza (selectbox + przycisk "Wczytaj")
# - Domieszki traktowane jak normalne składniki (udział objętościowy, bilans, koszt, CO2)
# - Odporne parsowanie liczb z przecinkiem i odstępami
# - Stabilne ID w edytorach (brak utraty fokusa)
# - Kruszywo: brak autokorekty inputów frakcji; normalizacja tylko do obliczeń
# - LIVE bez "zatwierdź": tabele nie są przebudowywane co rerun, tylko gdy zmieni się wybór
# - Edycja liczb jako tekst (żeby "12,5" nie robiło NaN->0 i nie powodowało cofek)
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
    """Odporne parsowanie pojedynczej wartości (np. z data_editor),
    obsługa przecinka i spacji, zwraca float lub nan."""
    if x is None:
        return math.nan
    if isinstance(x, (int, float)) and not (isinstance(x, float) and math.isnan(x)):
        return float(x)
    s = (
        str(x)
        .replace("\xa0", " ")
        .replace(" ", "")
        .replace(",", ".")
        .strip()
    )
    return pd.to_numeric(s, errors="coerce")


def fmt_num(x: Any, nd: int = 3) -> str:
    """Format do trzymania w polu tekstowym edycji."""
    try:
        v = float(x)
        if math.isnan(v):
            return ""
        # bez naukowego
        return f"{v:.{nd}f}".rstrip("0").rstrip(".")
    except Exception:
        return ""


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

    # kolumny tekstowe do edycji (ważne: edytujemy tekst, nie float)
    if "udzial_txt" not in t.columns:
        t["udzial_txt"] = t["udzial_pct"].apply(lambda x: fmt_num(x, 3))

    if "share_in_agg_pct" not in t.columns:
        # dla nie-kruszywa nieistotne, ale zostawiamy opcjonalnie
        pass
    if "share_in_agg_txt" not in t.columns:
        if "share_in_agg_pct" in t.columns:
            t["share_in_agg_txt"] = t["share_in_agg_pct"].apply(lambda x: fmt_num(x, 3))
        else:
            t["share_in_agg_txt"] = ""

    t["id"] = pd.to_numeric(t["id"], errors="coerce")
    t["gestosc_kgm3"] = pd.to_numeric(t["gestosc_kgm3"], errors="coerce")
    # uwaga: udzial_pct będziemy liczyć dynamicznie z udzial_txt/share_in_agg_txt

    tech_cols = ["id", "nazwa", "kategoria", "gestosc_kgm3", "udzial_pct", "udzial_txt", "share_in_agg_txt"]
    rest = [c for c in t.columns if c not in tech_cols]
    return t[tech_cols + rest]


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


def build_table_for_selection(df_all: pd.DataFrame, ids: List[int]) -> pd.DataFrame:
    ids = [int(x) for x in ids]
    if not ids:
        return pd.DataFrame(columns=["id", "nazwa", "kategoria", "gestosc_kgm3", "udzial_pct", "udzial_txt", "share_in_agg_txt"])
    t = df_all[df_all["id"].isin(ids)].copy()
    t = ensure_table_columns(t)
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
    # reset selekcji i tabel
    for cat_key, _ in CATEGORIES_ORDERED:
        st.session_state[f"msel_{cat_key}"] = []
        st.session_state.pop(f"tbl_{cat_key}", None)
        st.session_state.pop(f"ids_{cat_key}", None)

    existing_ids = set(df_all["id"].astype(int).tolist())

    # total kruszywa z receptury (sum udzial_pct kruszywo)
    kr_total = float(
        pd.to_numeric(df_loaded.loc[df_loaded["kategoria"].astype(str).str.strip().eq("kruszywo"), "udzial_pct"], errors="coerce")
        .fillna(0.0)
        .sum()
    )
    st.session_state["kruszywo_total_txt"] = fmt_num(kr_total, 3)

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

        ids = sdf["id"].tolist()
        st.session_state[f"msel_{cat_key}"] = ids

        base_tbl = build_table_for_selection(df_all, ids)
        base_tbl = ensure_table_columns(base_tbl)

        ud_map = dict(zip(
            sdf["id"].astype(int).tolist(),
            pd.to_numeric(sdf["udzial_pct"], errors="coerce").fillna(0.0).tolist()
        ))

        if cat_key == "kruszywo":
            # share txt jako procent w kruszywie (z udzial_pct / kr_total)
            if kr_total > 0:
                base_tbl["share_in_agg_txt"] = base_tbl["id"].map(lambda i: fmt_num((ud_map.get(int(i), 0.0) / kr_total) * 100.0, 3))
            else:
                n = max(len(base_tbl), 1)
                base_tbl["share_in_agg_txt"] = [fmt_num(100.0 / n, 3)] * len(base_tbl)

            # udzial_txt nieużywany dla kruszywa (liczymy z share + total), ale zostawmy spójnie
            base_tbl["udzial_txt"] = base_tbl["id"].map(lambda i: fmt_num(ud_map.get(int(i), 0.0), 3))
        else:
            base_tbl["udzial_txt"] = base_tbl["id"].map(lambda i: fmt_num(ud_map.get(int(i), 0.0), 3))

        st.session_state[f"tbl_{cat_key}"] = base_tbl.reset_index(drop=True).copy()
        st.session_state[f"ids_{cat_key}"] = tuple(sorted(ids))


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
# LEWA KOLUMNA: wybór + edycja udziałów (LIVE, bez cofek)
# ============================================================
with col_left:
    st.markdown("## Wybór i udziały składników")

    # init kruszywo total jako tekst
    if "kruszywo_total_txt" not in st.session_state:
        st.session_state["kruszywo_total_txt"] = "0"

    for cat_key, cat_title in CATEGORIES_ORDERED:
        st.subheader(cat_title)

        sel_key = f"msel_{cat_key}"
        init_state(sel_key, [])

        opts = options_for_category(df_all, cat_key)

        current_sel = st.session_state.get(sel_key, [])
        # sanitizacja
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
        sel_ids_tuple = tuple(sorted([int(x) for x in sel]))

        skey = f"tbl_{cat_key}"
        prev_ids = st.session_state.get(f"ids_{cat_key}", tuple())

        # REBUILD TABELI tylko jeśli zmienił się wybór (to mocno redukuje cofki)
        if sel_ids_tuple != prev_ids:
            base_tbl = build_table_for_selection(df_all, list(sel_ids_tuple))
            base_tbl = ensure_table_columns(base_tbl)

            # zachowaj edycje z poprzedniej tabeli, jeśli były
            if skey in st.session_state and isinstance(st.session_state[skey], pd.DataFrame) and not st.session_state[skey].empty:
                old = ensure_table_columns(st.session_state[skey].copy())
                old["id"] = pd.to_numeric(old["id"], errors="coerce")
                old = old.dropna(subset=["id"]).copy()
                old["id"] = old["id"].astype(int)
                old = old.drop_duplicates(subset=["id"], keep="first").set_index("id")

                # tekstowe udziały
                if "udzial_txt" in old.columns:
                    m = old["udzial_txt"].to_dict()
                    base_tbl["udzial_txt"] = base_tbl["id"].astype(int).map(m).combine_first(base_tbl["udzial_txt"])

                if cat_key == "kruszywo" and "share_in_agg_txt" in old.columns:
                    m2 = old["share_in_agg_txt"].to_dict()
                    base_tbl["share_in_agg_txt"] = base_tbl["id"].astype(int).map(m2).combine_first(base_tbl["share_in_agg_txt"])

            # init kruszywo split jeśli puste
            if cat_key == "kruszywo":
                # jeśli wczytane udzial_txt ma jakieś wartości, spróbuj z nich zrobić share
                if (base_tbl["share_in_agg_txt"].astype(str).str.strip() == "").all():
                    # z udzial_txt lub udzial_pct
                    ud = base_tbl["udzial_txt"].apply(to_num_val).fillna(0.0)
                    s = float(ud.sum())
                    if s > 0:
                        base_tbl["share_in_agg_txt"] = (ud / s * 100.0).apply(lambda v: fmt_num(v, 3))
                    else:
                        n = max(len(base_tbl), 1)
                        base_tbl["share_in_agg_txt"] = [fmt_num(100.0 / n, 3)] * len(base_tbl)

                # jeśli total nieustawiony, ustaw na sumę udzial_txt
                if "kruszywo_total_txt" not in st.session_state or str(st.session_state["kruszywo_total_txt"]).strip() == "":
                    st.session_state["kruszywo_total_txt"] = fmt_num(float(base_tbl["udzial_txt"].apply(to_num_val).fillna(0.0).sum()), 3)

            st.session_state[skey] = base_tbl.reset_index(drop=True).copy()
            st.session_state[f"ids_{cat_key}"] = sel_ids_tuple

        tbl = ensure_table_columns(st.session_state.get(skey, pd.DataFrame()).copy())
        if tbl.empty:
            continue

        # --- kruszywo: total + split (LIVE, edycja tekstowa) ---
        if cat_key == "kruszywo":
            st.text_input(
                "Kruszywo (razem) — Udział objętościowy [%]",
                key="kruszywo_total_txt",
                label_visibility="visible",
            )

            split_view = (
                tbl[["id", "nazwa", "share_in_agg_txt"]]
                .rename(columns={"share_in_agg_txt": "Udział w kruszywie [%]"})
                .set_index("id", drop=False)
            )

            edited_split = st.data_editor(
                split_view,
                key="editor_kruszywo_split_txt",
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                    "nazwa": st.column_config.TextColumn("nazwa", disabled=True),
                    "Udział w kruszywie [%]": st.column_config.TextColumn("Udział w kruszywie [%]"),
                },
            )

            if edited_split is not None and not edited_split.empty:
                # update tylko kolumny tekstowej, bez żadnych normalize/astype
                ed = edited_split.reset_index(drop=True).copy()
                ed = ed.rename(columns={"Udział w kruszywie [%]": "share_in_agg_txt"})
                # mapowanie po ID
                m = dict(zip(ed["id"].astype(int).tolist(), ed["share_in_agg_txt"].astype(str).tolist()))
                tbl["share_in_agg_txt"] = tbl["id"].astype(int).map(m).fillna(tbl["share_in_agg_txt"])
                st.session_state[skey] = tbl.copy()

            continue

        # --- inne kategorie (w tym domieszka): edycja tekstowa udziału obj. ---
        df_for_edit = (
            tbl[["id", "nazwa", "udzial_txt"]]
            .rename(columns={"udzial_txt": "Udział objętościowy [%]"})
            .set_index("id", drop=False)
        )

        edited = st.data_editor(
            df_for_edit,
            key=f"editor_{cat_key}_volume_txt",
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "nazwa": st.column_config.TextColumn("nazwa", disabled=True),
                "Udział objętościowy [%]": st.column_config.TextColumn("Udział objętościowy [%]"),
            },
        )

        if edited is not None and not edited.empty:
            ed = edited.reset_index(drop=True).copy()
            ed = ed.rename(columns={"Udział objętościowy [%]": "udzial_txt"})
            m = dict(zip(ed["id"].astype(int).tolist(), ed["udzial_txt"].astype(str).tolist()))
            tbl["udzial_txt"] = tbl["id"].astype(int).map(m).fillna(tbl["udzial_txt"])
            st.session_state[skey] = tbl.copy()


# ============================================================
# PRAWA KOLUMNA: bilans, tabele, parametry, zapis, wykresy
# ============================================================
with col_right:
    parts_base = []

    # parse total kruszywa
    kr_total_pct = float(to_num_val(st.session_state.get("kruszywo_total_txt", "0")))
    if math.isnan(kr_total_pct):
        kr_total_pct = 0.0

    for cat_key, _ in CATEGORIES_ORDERED:
        skey = f"tbl_{cat_key}"
        if skey in st.session_state and isinstance(st.session_state[skey], pd.DataFrame):
            df = ensure_table_columns(st.session_state[skey].copy())
            if df.empty:
                continue

            df["kategoria"] = cat_key

            # WYLICZ udzial_pct z tekstu (LIVE, bez cofek)
            if cat_key == "kruszywo":
                shares = df.get("share_in_agg_txt", "").astype(str).apply(to_num_val).fillna(0.0)
                share_sum = float(shares.sum())
                if share_sum > 0:
                    df["udzial_pct"] = (shares / share_sum) * kr_total_pct
                else:
                    df["udzial_pct"] = 0.0
            else:
                df["udzial_pct"] = df.get("udzial_txt", "").astype(str).apply(to_num_val).fillna(0.0)

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
