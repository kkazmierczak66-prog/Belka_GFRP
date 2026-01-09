import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import math

from io import BytesIO
from datetime import datetime
import matplotlib.pyplot as plt

# ==========================
# Google Sheets – receptury
# ==========================
GS_RECIPES_READY = False
try:
    from google.oauth2.service_account import Credentials
    import gspread
    try:
        from gspread_dataframe import get_as_dataframe
    except Exception:
        get_as_dataframe = None

    GSA = "gcp_service_account"
    SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")
    SHEET_RECIPES = st.secrets.get("SHEET_RECIPES", "receptury")
    SHEET_MATERIALS = st.secrets.get("SHEET_MATERIALS", "materiały")
    if GSA in st.secrets and SPREADSHEET_ID:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        CREDS = Credentials.from_service_account_info(st.secrets[GSA], scopes=SCOPES)
        GS_RECIPES_READY = True
except Exception:
    GS_RECIPES_READY = False
    get_as_dataframe = None

# --- Arkusz z prętami GFRP ---
SHEET_GFRP = st.secrets.get("SHEET_GFRP", "gfrp_bars")


@st.cache_data(show_spinner=False)
def read_gfrp_bars(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """
    Czyta bazę prętów GFRP z Google Sheets (arkusz SHEET_GFRP).
    Zwraca kolumny:
    id, nazwa, srednica_mm, profil, R_t_MPa, E_GPa, τ_base_MPa,
    gestosc_gcm3, cena_pln, cena_za, co2e_kgkg
    """
    if not GS_RECIPES_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account).")

    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    try:
        ws = ss.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        raise RuntimeError(f"Nie znaleziono zakładki „{sheet_name}” w pliku Google Sheets.")

    if get_as_dataframe is not None:
        df = get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
    else:
        rows = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(rows)

    wanted = [
        "id", "nazwa", "srednica_mm", "profil",
        "R_t_MPa", "E_GPa", "τ_base_MPa", "gestosc_gcm3",
        "cena_pln", "cena_za", "co2e_kgkg",
    ]
    for c in wanted:
        if c not in df.columns:
            df[c] = None

    # liczby
    for c in ["id", "srednica_mm", "R_t_MPa", "E_GPa", "τ_base_MPa", "gestosc_gcm3", "cena_pln", "co2e_kgkg"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["profil"] = df["profil"].astype(object)
    df["cena_za"] = df["cena_za"].astype(object)

    return df[wanted]


@st.cache_data(show_spinner=False)
def read_recipes_summary(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """
    Czyta arkusz receptur i zwraca wiersze PODSUMOWAŃ (SUMMARY / __SUMMARY__),
    zwracając DataFrame z kolumnami:
    recipe_name, gestosc_mix_kgm3, fck_mpa, fctm_mpa, ecm_gpa

    Wersja odporna na różne warianty nagłówków (spacje, wielkość liter, aliasy).
    """
    if not GS_RECIPES_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account).")

    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    try:
        ws = ss.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        raise RuntimeError(f"Nie znaleziono zakładki „{sheet_name}” w pliku Google Sheets.")

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(columns=["recipe_name", "gestosc_mix_kgm3", "fck_mpa", "fctm_mpa", "ecm_gpa"])

    # Domyślne indeksy (zgodne z typowym HEADERS):
    # timestamp, recipe_name, material_id, nazwa, kategoria,
    # gestosc_kgm3, udzial_pct, obj_m3, masa_kgm3, sum_obj_m3m3, sum_mas_kgm3,
    # gestosc_mix_kgm3, w_c, fck_mpa, fctm_mpa, ecm_gpa
    default_idx = {
        "recipe_name": 1,        # B
        "nazwa": 3,              # D
        "gestosc_mix_kgm3": 11,  # L
        "fck_mpa": 13,           # N
        "fctm_mpa": 14,          # O
        "ecm_gpa": 15,           # P
    }

    header = values[0] if values else []

    def norm(s: str) -> str:
        return " ".join((s or "").strip().lower().split())

    header_norm = [norm(h) for h in header]
    looks_like_header = ("recipe_name" in header) or ("timestamp" in header) or ("recipe name" in header_norm)

    # aliasy dla kolumn po normalizacji
    aliases = {
        "recipe_name": {"recipe_name", "recipe name", "nazwa receptury"},
        "nazwa": {"nazwa", "name"},
        "gestosc_mix_kgm3": {
            "gestosc_mix_kgm3", "gęstość mieszanki [kg/m3]", "gestosc mieszanki [kg/m3]",
            "gestosc mix kgm3", "gestosc mieszanki kg/m3", "rho_mix", "rho mix", "rho [kg/m3]"
        },
        "fck_mpa": {"fck_mpa", "fck [mpa]", "fck", "fck mpa"},
        "fctm_mpa": {"fctm_mpa", "fctm [mpa]", "fctm", "fctm mpa"},
        "ecm_gpa": {"ecm_gpa", "ecm [gpa]", "ecm", "e_c", "e_c [gpa]", "ecm gpa", "ecm (gpa)"},
    }

    def idx(col_key: str) -> int:
        if header:
            targets = aliases.get(col_key, {col_key})
            for i, h in enumerate(header_norm):
                if h in targets:
                    return i
        return default_idx[col_key]

    i_recipe = idx("recipe_name")
    i_nazwa = idx("nazwa")
    i_rho_mix = idx("gestosc_mix_kgm3")
    i_fck = idx("fck_mpa")
    i_fctm = idx("fctm_mpa")
    i_ecm = idx("ecm_gpa")

    data_rows = values[1:] if looks_like_header else values
    records = []
    max_i = max(i_recipe, i_nazwa, i_rho_mix, i_fck, i_fctm, i_ecm)

    def to_float(x):
        s = str(x).strip()
        if not s:
            return float("nan")
        s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return float("nan")

    for row in data_rows:
        r = list(row) + [""] * (max_i + 1 - len(row))
        marker = str(r[i_nazwa]).strip().upper()
        if marker not in ("SUMMARY", "__SUMMARY__"):
            continue

        recipe_name = str(r[i_recipe]).strip()
        records.append({
            "recipe_name": recipe_name,
            "gestosc_mix_kgm3": to_float(r[i_rho_mix]),
            "fck_mpa": to_float(r[i_fck]),
            "fctm_mpa": to_float(r[i_fctm]),
            "ecm_gpa": to_float(r[i_ecm]),
        })

    if not records:
        return pd.DataFrame(columns=["recipe_name", "gestosc_mix_kgm3", "fck_mpa", "fctm_mpa", "ecm_gpa"])

    return pd.DataFrame(records)

# ============================================================
# WCZYTYWANIE BELKI Z GOOGLE SHEETS (arkusz: "belki tpd")
# ============================================================

SHEET_BEAMS_TPD = st.secrets.get("SHEET_BEAMS_TPD", "belki tpd")

def _to_float(x, default=None):
    try:
        if x is None:
            return default
        s = str(x).strip()
        if not s:
            return default
        s = s.replace(",", ".")
        return float(s)
    except Exception:
        return default

def _to_int(x, default=None):
    try:
        if x is None:
            return default
        s = str(x).strip()
        if not s:
            return default
        s = s.replace(",", ".")
        return int(float(s))
    except Exception:
        return default


@st.cache_data(show_spinner=False)
def read_beams_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    if not GS_RECIPES_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account).")

    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    ws = ss.worksheet(sheet_name)

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    header = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)

    # pomocniczo: sortowanie po ID
    if "ID" in df.columns:
        df["_ID_num"] = pd.to_numeric(
            df["ID"].astype(str).str.replace(",", ".", regex=False),
            errors="coerce"
        )
    else:
        df["_ID_num"] = pd.NA

    return df


def apply_loaded_beam_tpd_to_state(row: dict):
    """
    Mapowanie 1:1 z kolumn INPUT_* (jak podałeś) na klucze widgetów.
    """
    # --- nazwa -> pod zapis/nadpisanie
    beam_name = str(row.get("Nazwa belki", "")).strip()
    if beam_name:
        st.session_state["beam_name_to_save"] = beam_name
        # (opcjonalnie) od razu ustaw checkbox nadpisania
        st.session_state["chk_overwrite_beam"] = True

    # --- beton
    beton_mode = str(row.get("INPUT_beton_mode", "")).strip() or "Wybór z bazy danych"
    st.session_state["beton_mode"] = beton_mode

    recipe_name = str(row.get("INPUT_beton_recipe_name", "")).strip()
    if recipe_name:
        st.session_state["beton_recipe_name"] = recipe_name

    # --- geometria TPD
    st.session_state["L_beam"]      = _to_float(row.get("INPUT_L_m"), 1.0)
    st.session_state["b_polki"]     = _to_float(row.get("INPUT_bf_cm"), 20.0)
    st.session_state["h_polki"]     = _to_float(row.get("INPUT_hf_cm"), 5.0)
    st.session_state["b_srodnika"]  = _to_float(row.get("INPUT_bw_cm"), 5.0)
    st.session_state["h_srodnika"]  = _to_float(row.get("INPUT_hw_cm"), 15.0)

    # --- ograniczenia masy
    st.session_state["masa_min"] = _to_float(row.get("INPUT_masa_min_kg"), 0.0)
    st.session_state["masa_max"] = _to_float(row.get("INPUT_masa_max_kg"), 0.0)

    # --- zbrojenie dolne (półka)
    st.session_state["z_ot_dolna"]       = _to_float(row.get("INPUT_z_ot_dolna_mm"), 5.0)
    st.session_state["z_ot_gorna"]       = _to_float(row.get("INPUT_z_ot_gorna_mm"), 5.0)
    st.session_state["z_ot_boczna"]      = _to_float(row.get("INPUT_z_ot_boczna_mm"), 5.0)
    st.session_state["z_odst_poziomy"]   = _to_float(row.get("INPUT_z_odst_poziomy_mm"), 5.0)
    st.session_state["z_odst_pionowy"]   = _to_float(row.get("INPUT_z_odst_pionowy_mm"), 5.0)
    st.session_state["z_n_wlasne"]       = _to_int(row.get("INPUT_z_n_wlasne"), 4)
    st.session_state["z_warstwy_wlasne"] = _to_int(row.get("INPUT_z_warstwy_wlasne"), 1)

    # --- zbrojenie górne (środnik)
    st.session_state["s_ot_gorna"]       = _to_float(row.get("INPUT_s_ot_gorna_mm"), 5.0)
    st.session_state["s_ot_dolna"]       = _to_float(row.get("INPUT_s_ot_dolna_mm"), 5.0)
    st.session_state["s_ot_boczna"]      = _to_float(row.get("INPUT_s_ot_boczna_mm"), 5.0)
    st.session_state["s_odst_poziomy"]   = _to_float(row.get("INPUT_s_odst_poziomy_mm"), 5.0)
    st.session_state["s_odst_pionowy"]   = _to_float(row.get("INPUT_s_odst_pionowy_mm"), 5.0)
    st.session_state["s_n_wlasne"]       = _to_int(row.get("INPUT_s_n_wlasne"), 2)
    st.session_state["s_warstwy_wlasne"] = _to_int(row.get("INPUT_s_warstwy_wlasne"), 1)

    # --- ścinanie
    shear_mode = str(row.get("INPUT_shear_choice_mode", "")).strip()
    if shear_mode:
        st.session_state["shear_choice_mode"] = shear_mode

    p_custom = _to_float(row.get("INPUT_shear_P_custom_kN"), None)
    if p_custom is not None and p_custom > 0:
        st.session_state["shear_P_custom_kN"] = float(p_custom)

    # --- GFRP bar: zapisujemy jako "pending", bo df_gfrp może jeszcze nie być wczytane
    gfrp_id = _to_int(row.get("INPUT_gfrp_bar_id"), None)
    if gfrp_id is not None:
        st.session_state["__pending_gfrp_bar_id"] = int(gfrp_id)


def resolve_pending_gfrp_selection(df_gfrp: pd.DataFrame):
    """
    Po wczytaniu bazy prętów mapuje __pending_gfrp_bar_id -> rebar_gfrp_sel (index selectboxa).
    """
    pending = st.session_state.get("__pending_gfrp_bar_id", None)
    if pending is None:
        return
    if df_gfrp is None or df_gfrp.empty or "id" not in df_gfrp.columns:
        return

    try:
        pending = int(pending)
    except Exception:
        return

    matches = df_gfrp.index[df_gfrp["id"].fillna(-1).astype(int) == pending].tolist()
    if matches:
        st.session_state["rebar_gfrp_sel"] = matches[0]

    # wyczyść pending
    del st.session_state["__pending_gfrp_bar_id"]


def load_beam_ui_tpd(sheet_name: str = SHEET_BEAMS_TPD):
    """
    UI: wybór belki z listy + przycisk "Wczytaj".
    Po wczytaniu resetuje wybór bez łamania zasad Streamlita (flaga + rerun).
    """
    st.subheader("Wczytaj belkę z bazy")

    try:
        df = read_beams_sheet(SPREADSHEET_ID, sheet_name)
    except Exception as e:
        st.error(f"Nie udało się wczytać arkusza „{sheet_name}”: {e}")
        return

    if df is None or df.empty:
        st.info(f"Arkusz „{sheet_name}” jest pusty.")
        return

    # sort po ID rosnąco
    if "_ID_num" in df.columns:
        df = df.sort_values("_ID_num", ascending=True)

    records = df.to_dict(orient="records")

    placeholder = "— wybierz belkę —"
    key_sel = "beam_tpd_to_load"
    key_reset = "__reset_beam_tpd_to_load"

    # ✅ reset MUSI być przed instancjacją selectboxa
    if st.session_state.get(key_reset, False):
        st.session_state[key_sel] = placeholder
        st.session_state[key_reset] = False

    if key_sel not in st.session_state:
        st.session_state[key_sel] = placeholder

    options = [placeholder] + [
        f"ID={r.get('ID','?')} • {r.get('Nazwa belki','(bez nazwy)')}"
        for r in records
    ]

    colA, colB = st.columns([6, 1.6], vertical_alignment="top")

    with colA:
        choice = st.selectbox("", options=options, key=key_sel)

    with colB:
        # spacer żeby przycisk był na wysokości pola selectboxa (a nie etykiety)
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        clicked = st.button("📥 Wczytaj", use_container_width=True, disabled=(choice == placeholder))

    if clicked and choice != placeholder:
        idx = options.index(choice) - 1
        row = records[idx]

        apply_loaded_beam_tpd_to_state(row)

        # ❌ nie ruszamy key_sel tutaj
        # ✅ ustawiamy flagę i robimy rerun
        st.session_state[key_reset] = True

        st.success(f"Wczytano belkę: {row.get('Nazwa belki','')}")
        st.rerun()


# ---------------------------
# USTAWIENIA STRONY
# ---------------------------
st.set_page_config(page_title="Belka GFRP", layout="wide")
# --- Tytuł + przycisk odświeżenia w jednej linii ---
title_col, refresh_col = st.columns([10, 1.7])     # ← tu kontrolujesz szerokość przycisku

with title_col:
    st.title("Definicja belki i obliczenia")

with refresh_col:
    if st.button("↻ Odśwież dane", use_container_width=True,
                 help="Wymusza ponowne pobranie danych z Google Sheets"):
        st.cache_data.clear()
        for key in ["beton_dane", "rebar_bar", "beton_mode", "beton_recipe_name"]:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

load_beam_ui_tpd(sheet_name=SHEET_BEAMS_TPD)

st.markdown("---")

# ============================================================
# SEKCJA – DANE DOTYCZĄCE BETONU
# ============================================================
st.header("Dane dotyczące betonu")

col_b1, col_b2 = st.columns([1, 2])

with col_b1:
    # --- FIX: st.radio(key="beton_mode") nie przyjmie wartości spoza options ---
    BETON_OPTIONS = ["Wybór z bazy danych", "Definiuj ręcznie (brak obliczeń punktacji)"]

    if "beton_mode" in st.session_state:
        v = st.session_state.get("beton_mode")
        if v not in BETON_OPTIONS:
            v_norm = str(v).strip().lower()
            if v_norm in ("gsheet", "google", "baza", "wybór z bazy danych", "wybor z bazy danych"):
                st.session_state["beton_mode"] = BETON_OPTIONS[0]
            elif v_norm in ("manual", "recznie", "ręcznie", "definiuj ręcznie (brak obliczeń punktacji)"):
                st.session_state["beton_mode"] = BETON_OPTIONS[1]
            else:
                st.session_state["beton_mode"] = BETON_OPTIONS[0]
    else:
        # opcjonalnie: ustaw domyślną wartość, jeśli klucz nie istnieje
        st.session_state["beton_mode"] = BETON_OPTIONS[0]

    beton_mode = st.radio(
        "Sposób wprowadzania danych:",
        ["Wybór z bazy danych", "Definiuj ręcznie (brak obliczeń punktacji)"],
        key="beton_mode"
    )

with col_b2:
    if beton_mode == "Wybór z bazy danych":
        # TYLKO Google Sheets (bez „Klas betonu”)
        if not GS_RECIPES_READY:
            st.error("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account).")
            st.session_state["beton_dane"] = {}
        else:
            try:
                df_recipes = read_recipes_summary(SPREADSHEET_ID, SHEET_RECIPES)
            except Exception as e:
                st.error(f"Nie udało się wczytać receptur z Google Sheets: {e}")
                df_recipes = pd.DataFrame()

            if df_recipes.empty:
                st.warning("Nie znaleziono żadnych zapisanych receptur (__SUMMARY__) w arkuszu receptury.")
                st.session_state["beton_dane"] = {}
            else:
                recipe_names = sorted(
                    df_recipes["recipe_name"].dropna().unique().tolist()
                )
                sel_recipe = st.selectbox(
                    "Receptura z bazy danych:",
                    recipe_names,
                    key="beton_recipe_name",
                )

                rec_row = (
                    df_recipes[df_recipes["recipe_name"] == sel_recipe]
                    .sort_index()
                    .iloc[-1]
                )

                rho_mix = float(rec_row["gestosc_mix_kgm3"]) if pd.notna(rec_row["gestosc_mix_kgm3"]) else math.nan
                f_ck = float(rec_row["fck_mpa"]) if pd.notna(rec_row["fck_mpa"]) else math.nan
                f_ctm = float(rec_row["fctm_mpa"]) if pd.notna(rec_row["fctm_mpa"]) else math.nan
                E_c = float(rec_row["ecm_gpa"]) if pd.notna(rec_row["ecm_gpa"]) else math.nan

                st.session_state["beton_dane"] = {
                    "source": "gsheet",
                    "klasa": sel_recipe,
                    "f_ck": f_ck,
                    "f_ctm": f_ctm,
                    "rho": rho_mix,
                    "E_c_GPa": E_c,
                }

                st.markdown(
                    f"""
                    - f_ck = {f_ck if not math.isnan(f_ck) else "—"} MPa  
                    - f_ctm = {f_ctm if not math.isnan(f_ctm) else "—"} MPa  
                    - ρ (gęstość mieszanki) = {rho_mix if not math.isnan(rho_mix) else "—"} kg/m³  
                    - E_c (Ecm) = {E_c if not math.isnan(E_c) else "—"} GPa  
                    """
                )

    else:  # "Definiuj ręcznie (brak obliczeń punktacji)"
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            f_ck_manual = st.number_input(
                "f_ck [MPa]",
                min_value=5.0,
                value=25.0,
                step=1.0,
                key="beton_fck"
            )
        with col_m2:
            f_ctm_manual = st.number_input(
                "f_ctm [MPa]",
                min_value=1.0,
                value=2.6,
                step=0.1,
                key="beton_fctm"
            )
        with col_m3:
            rho_manual = st.number_input(
                "ρ mieszanki [kg/m³]",
                min_value=1000.0,
                value=2400.0,
                step=50.0,
                key="beton_rho"
            )
        with col_m4:
            E_c_manual = st.number_input(
                "E_c [GPa]",
                min_value=10.0,
                value=30.0,
                step=1.0,
                key="beton_Ec"
            )

        st.session_state["beton_dane"] = {
            "source": "manual",
            "klasa": None,
            "f_ck": f_ck_manual,
            "f_ctm": f_ctm_manual,
            "rho": rho_manual,
            "E_c_GPa": E_c_manual,
        }






st.markdown("---")
# --- DEFAULTY do widgetów (żeby nie było warningów o value + session_state) ---
TPD_DEFAULTS = {
    "tpd_L_beam": 1.0,
    "tpd_b_polki": 20.0,
    "tpd_h_polki": 5.0,
    "tpd_b_srodnika": 5.0,
    "tpd_h_srodnika": 15.0,
    "tpd_masa_min": 5.0,
    "tpd_masa_max": 15.0,
}
for k, v in TPD_DEFAULTS.items():
    st.session_state.setdefault(k, v)
# ---------------------------
# PARAMETRY BELKI
# ---------------------------
beton_info = st.session_state.get("beton_dane", {}) or {}
rho = beton_info.get("rho", 2400.0)
try:
    if rho is None or math.isnan(float(rho)):
        rho = 2400.0
    else:
        rho = float(rho)
except Exception:
    rho = 2400.0



st.header("Parametry belki")

st.header("Parametry belki")

# --- 1) Parametry geometryczne ---
st.subheader("Parametry geometryczne")

col_geom1, col_geom2, col_geom3, col_geom4 = st.columns(4)
with col_geom1:
    L_beam = st.number_input(
        "Długość belki L [m]",
        min_value=0.01,
        step=0.1,
        key="tpd_L_beam",
    )
with col_geom2:
    b_polki = st.number_input(
        "Szerokość półki b_f [cm]",
        min_value=0.0,
        step=0.5,
        key="tpd_b_polki",
    )
with col_geom3:
    h_polki = st.number_input(
        "Wysokość półki h_f [cm]",
        min_value=0.0,
        step=0.5,
        key="tpd_h_polki",
    )
with col_geom4:
    b_srodnika = st.number_input(
        "Szerokość środnika b_w [cm]",
        min_value=0.0,
        step=0.5,
        key="tpd_b_srodnika",
    )

col_geom5, col_geom6 = st.columns(2)
with col_geom5:
    h_srodnika = st.number_input(
        "Wysokość środnika h_w [cm]",
        min_value=0.0,
        step=0.5,
        key="tpd_h_srodnika",
    )
with col_geom6:
    st.markdown("<br>", unsafe_allow_html=True)

# --- 2) Ograniczenia ---
st.subheader("Ograniczenia")

col_lim1, col_lim2 = st.columns(2)
with col_lim1:
    masa_min = st.number_input(
        "Minimalna masa belki [kg]",
        min_value=0.0,
        step=0.5,
        key="tpd_masa_min",
    )
with col_lim2:
    masa_max = st.number_input(
        "Maksymalna masa belki [kg]",
        min_value=0.0,
        step=0.5,
        key="tpd_masa_max",
    )





st.markdown("---")

# ---------------------------
# OBLICZENIA GEOMETRII
# ---------------------------
b_polki_m = b_polki / 100.0
h_polki_m = h_polki / 100.0
b_srodnika_m = b_srodnika / 100.0
h_srodnika_m = h_srodnika / 100.0

A_min = masa_min / (rho * L_beam) if rho > 0 and L_beam > 0 else 0.0
A_max = masa_max / (rho * L_beam) if rho > 0 and L_beam > 0 else 0.0

A_polka = b_polki_m * h_polki_m
A_srodnik = b_srodnika_m * h_srodnika_m
A = A_polka + A_srodnik

if A > 0:
    y_c = (A_polka * (h_polki_m / 2.0) + A_srodnik * (h_polki_m + h_srodnika_m / 2.0)) / A
else:
    y_c = 0.0

I = (
    (b_polki_m * h_polki_m ** 3) / 12.0 + (A_polka * (y_c - h_polki_m / 2.0) ** 2 if A > 0 else 0.0)
    + (b_srodnika_m * h_srodnika_m ** 3) / 12.0 + (
        A_srodnik * (h_polki_m + h_srodnika_m / 2.0 - y_c) ** 2 if A > 0 else 0.0)
)
I_do_A = I / A if A > 0 else 0.0

A_cm2, A_min_cm2, A_max_cm2 = A * 1e4, A_min * 1e4, A_max * 1e4
I_cm4, I_do_A_cm2, y_c_cm = I * 1e8, I_do_A * 1e4, y_c * 100.0

# --- MASA BELKI (kg) ---
masa_belki = rho * A * L_beam  # [kg]

# ---------------------------
# WALIDACJA
# ---------------------------
if A > 0 and A_min <= A <= A_max and A_min < A_max:
    status_text, status_color = "✅ Pole przekroju mieści się w zakresie z masy.", "lime"
elif A == 0:
    status_text, status_color = "⚠️ Pole przekroju wynosi 0. Zwiększ wymiary.", "orange"
elif A_min >= A_max:
    status_text, status_color = "⚠️ Zakres mas jest niespójny (masa_min ≥ masa_max).", "orange"
else:
    status_text, status_color = "❌ Pole przekroju poza zakresem wynikającym z masy.", "red"

# ---------------------------
# WYNIKI + WYKRES
# ---------------------------
col_l, col_r = st.columns([1, 1])

with col_l:
    st.subheader("Obliczenia - geometria")
    st.markdown(
        f"""
- Minimalne pole przekroju: **{A_min_cm2:.2f} cm²**  
- Maksymalne pole przekroju: **{A_max_cm2:.2f} cm²**  
- Pole przekroju **A**: **{A_cm2:.2f} cm²**  
- Moment bezwładności **I**: **{I_cm4:.2f} cm⁴**  
- Stosunek **I/A**: **{I_do_A_cm2:.2f} cm²**  
- Środek ciężkości **y_c**: **{y_c_cm:.2f} cm**  
- **Masa belki**: **{masa_belki:.2f} kg**
"""
    )
    st.markdown(f"<span style='color:{status_color}; font-weight:700'>{status_text}</span>", unsafe_allow_html=True)

    if masa_min < masa_max and masa_belki > 0:
        if masa_min <= masa_belki <= masa_max:
            st.markdown("<span style='color:lime; font-weight:700'>✅ Masa belki mieści się w zadanym zakresie.</span>", unsafe_allow_html=True)
        else:
            st.markdown("<span style='color:red; font-weight:700'>❌ Masa belki poza zadanym zakresem.</span>", unsafe_allow_html=True)

# ---------------------------
# WYKRES: AUTO-SKALOWANIE
# ---------------------------
H_cm = h_polki + h_srodnika
Bmax_cm = max(b_polki, b_srodnika)

pad_x = max(1.0, 0.1 * Bmax_cm)
pad_y = max(1.0, 0.1 * H_cm)

x_min, x_max = -Bmax_cm / 2.0 - pad_x, Bmax_cm / 2.0 + pad_x
y_min, y_max = -pad_y / 2.0, H_cm + pad_y

fig = go.Figure()
fig.update_layout(
    plot_bgcolor="black",
    paper_bgcolor="black",
    font=dict(color="white")
)

# ==============================
# RYSOWANIE PÓŁKI I ŚRODNIKA
# ==============================
if b_polki > 0 and h_polki > 0:
    fig.add_shape(
        type="rect",
        x0=-b_polki / 2,
        y0=0,
        x1=b_polki / 2,
        y1=h_polki,
        line=dict(color="white"),
        fillcolor="lightblue"
    )

if b_srodnika > 0 and h_srodnika > 0:
    fig.add_shape(
        type="rect",
        x0=-b_srodnika / 2,
        y0=h_polki,
        x1=b_srodnika / 2,
        y1=h_polki + h_srodnika,
        line=dict(color="white"),
        fillcolor="orange"
    )

# ==============================
# ŚRODEK CIĘŻKOŚCI
# ==============================
if A > 0:
    fig.add_trace(
        go.Scatter(
            x=[0], y=[y_c_cm],
            mode="markers",
            marker=dict(color="red", size=12, line=dict(color="white", width=2))
        )
    )

    fig.add_shape(
        type="line",
        x0=x_min, y0=y_c_cm,
        x1=x_max, y1=y_c_cm,
        line=dict(color="red", width=2, dash="dash")
    )

    fig.add_annotation(
        x=x_max + pad_x * 0.4,  # przesunięcie w prawo
        y=y_c_cm,
        text=f"y_c = {y_c_cm:.2f} cm",
        showarrow=False,
        font=dict(color="red"),
        xanchor="left"
    )


# ==================================================
# FUNKCJE POMOCNICZE DO LINII WYMIAROWYCH Z TICKAMI
# ==================================================
def add_dim_h(fig, x0, x1, y, tick=0.6, w=2):
    """Pozioma linia wymiarowa + pionowe ticki."""
    fig.add_shape(type="line", x0=x0, y0=y, x1=x1, y1=y, line=dict(color="white", width=w))
    fig.add_shape(type="line", x0=x0, y0=y - tick, x1=x0, y1=y + tick, line=dict(color="white", width=w))
    fig.add_shape(type="line", x0=x1, y0=y - tick, x1=x1, y1=y + tick, line=dict(color="white", width=w))


def add_dim_v(fig, x, y0, y1, tick=0.6, w=2):
    """Pionowa linia wymiarowa + poziome ticki."""
    fig.add_shape(type="line", x0=x, y0=y0, x1=x, y1=y1, line=dict(color="white", width=w))
    fig.add_shape(type="line", x0=x - tick, y0=y0, x1=x + tick, y1=y0, line=dict(color="white", width=w))
    fig.add_shape(type="line", x0=x - tick, y0=y1, x1=x + tick, y1=y1, line=dict(color="white", width=w))


# Automatyczna skala ticków
tick_h = max(0.4, 0.03 * Bmax_cm)
tick_v = max(0.4, 0.03 * H_cm)

# ==============================
# LINIE WYMIAROWE
# ==============================

# b_f – pozioma linia na dole
add_dim_h(fig, -b_polki / 2, b_polki / 2, y_min + 0.4, tick=tick_v)
fig.add_annotation(
    x=0, y=y_min,
    text=f"b_f = {b_polki} cm",
    showarrow=False,
    font=dict(color="white")
)

# b_w – pozioma linia na górze
add_dim_h(fig, -b_srodnika / 2, b_srodnika / 2, y_max - 0.4, tick=tick_v)
fig.add_annotation(
    x=0, y=y_max - 0.8,
    text=f"b_w = {b_srodnika} cm",
    showarrow=False,
    font=dict(color="white")
)
add_dim_v(fig, x_min + 0.5, 0, h_polki, tick=tick_h)
fig.add_annotation(
    x=x_min + 0.3 - 0.7,   # odsunięcie od linii wymiarowej
    y=h_polki / 2,
    text=f"h_f = {h_polki} cm",
    showarrow=False,
    textangle=90,
    font=dict(color="white")
)


# h_w – pionowa linia przy lewej krawędzi środnika
add_dim_v(fig, x_min + 1.5, h_polki, h_polki + h_srodnika, tick=tick_h)
fig.add_annotation(
    x=x_min + 1.3 - 0.7,
    y=h_polki + h_srodnika / 2,
    text=f"h_w = {h_srodnika} cm",
    showarrow=False,
    textangle=90,
    font=dict(color="white")
)

# H – pionowa linia po prawej stronie
add_dim_v(fig, x_max - 0.5, 0, H_cm, tick=tick_h)
fig.add_annotation(
    x=x_max - 0.7 + 0.7,   # odsuwamy w prawo
    y=H_cm / 2,
    text=f"H = {H_cm} cm",
    showarrow=False,
    textangle=90,
    font=dict(color="white")
)

# ==============================
# USTAWIENIA OSI I WYŚWIETLENIE
# ==============================
fig.update_layout(
    width=None,
    height=650,
    margin=dict(l=10, r=10, t=40, b=10),
    showlegend=False
)

fig.update_yaxes(
    scaleanchor="x", scaleratio=1,
    gridcolor="#444", zerolinecolor="#666"
)
fig.update_xaxes(
    range=[x_min, x_max], title="Szerokość [cm], oś x",
    gridcolor="#444", zerolinecolor="#666"
)
fig.update_yaxes(
    range=[y_min, y_max], title="Wysokość [cm], oś y"
)

with col_r:

    st.plotly_chart(fig, use_container_width=True)


# ---------------------------
# SEKCJA ZBROJENIE – pręt z bazy GFRP (średnica tylko z bazy)
# ---------------------------
st.markdown("---")
st.header("Zbrojenie")

# === 1) Wybór pręta z bazy GFRP ===


df_gfrp = None

if GS_RECIPES_READY:
    try:
        df_gfrp_tmp = read_gfrp_bars(SPREADSHEET_ID, SHEET_GFRP)
        if df_gfrp_tmp is not None and not df_gfrp_tmp.empty:
            df_gfrp = df_gfrp_tmp.copy()
    except Exception as e:
        st.error(f"Nie udało się wczytać bazy prętów GFRP z Google Sheets: {e}")
else:
    st.error("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account) – baza prętów niedostępna.")

if df_gfrp is not None and not df_gfrp.empty:
    resolve_pending_gfrp_selection(df_gfrp)

if df_gfrp is None or df_gfrp.empty:
    st.warning("Baza prętów GFRP jest pusta. Uzupełnij arkusz SHEET_GFRP, aby korzystać z tej sekcji.")
else:
    options_idx = df_gfrp.index.tolist()

    def format_bar(i: int) -> str:
        r = df_gfrp.loc[i]
        parts = []
        if pd.notna(r.get("id")):
            try:
                parts.append(f"{int(r['id'])}")
            except Exception:
                parts.append(str(r["id"]))
        name = str(r.get("nazwa", "") or "").strip()
        if name:
            parts.append(name)
        fi = r.get("srednica_mm")
        if pd.notna(fi):
            parts.append(f"⌀{fi:g} mm")
        prof = str(r.get("profil", "") or "").strip()
        if prof:
            parts.append(prof)
        return " – ".join(parts) if parts else str(i)

    sel_idx = st.selectbox(
        "Pręt z bazy danych:",
        options_idx,
        format_func=format_bar,
        key="rebar_gfrp_sel",
    )
    row = df_gfrp.loc[sel_idx]

    # ŚREDNICA tylko z bazy (wymagana)
    phi_mm = None
    if pd.notna(row.get("srednica_mm")):
        try:
            phi_mm = float(row["srednica_mm"])
        except Exception:
            phi_mm = None

    if not phi_mm or phi_mm <= 0:
        st.error("Wybrany pręt nie ma poprawnie zdefiniowanej średnicy 'srednica_mm'. Uzupełnij arkusz i spróbuj ponownie.")
        st.stop()

    rho_bar = float(row["gestosc_gcm3"]) * 1000.0 if pd.notna(row.get("gestosc_gcm3")) else float("nan")

    selected_bar_info = {
        "id": int(row["id"]) if pd.notna(row.get("id")) else None,
        "name": str(row.get("nazwa", "") or ""),
        "phi_mm": phi_mm,  # ← tylko z bazy
        "profil": str(row.get("profil", "") or ""),
        "R_t_MPa": float(row["R_t_MPa"]) if pd.notna(row.get("R_t_MPa")) else float("nan"),
        "E_GPa": float(row["E_GPa"]) if pd.notna(row.get("E_GPa")) else float("nan"),
        "tau_base_MPa": float(row["τ_base_MPa"]) if pd.notna(row.get("τ_base_MPa")) else float("nan"),
        "rho_kgm3": rho_bar,
        "price_pln": float(row["cena_pln"]) if pd.notna(row.get("cena_pln")) else float("nan"),
        "price_unit": str(row.get("cena_za", "") or ""),
        "co2e_kgkg": float(row["co2e_kgkg"]) if pd.notna(row.get("co2e_kgkg")) else float("nan"),

    }
    st.session_state["rebar_bar"] = selected_bar_info

    st.markdown(
        f"""
  
- ⌀ = {phi_mm:.1f} mm, profil = {selected_bar_info.get("profil") or "—"}  
- Rₜ ≈ {selected_bar_info.get("R_t_MPa", float("nan")):.0f} MPa, E ≈ {selected_bar_info.get("E_GPa", float("nan")):.1f} GPa  
- Gęstość pręta ≈ {rho_bar:.0f} kg/m³  
- Cena: {selected_bar_info.get("price_pln", float("nan")):.2f} USD / {selected_bar_info.get("price_unit") or "—"}  
- Ślad CO₂e ≈ {selected_bar_info.get("co2e_kgkg", float('nan')):.3f} kg/kg  
"""
    )

    # Ustal średnicę dla półki i środnika – bez inputów
    phi_f_mm = phi_mm  # półka
    phi_w_mm = phi_mm  # środnik



    # === 2) Rozmieszczenie prętów w półce (bez pola średnicy) ===
    st.subheader("Rozmieszczenie prętów w półce")

    col_left, col_right = st.columns([1, 1])

    with col_left:
        col1, col2, col3 = st.columns(3)
        with col1:
            otulina_dolna = st.number_input("Otulina dolna [mm]", value=5.0, key="z_ot_dolna")
        with col2:
            otulina_gorna = st.number_input("Otulina górna [mm]", value=5.0, key="z_ot_gorna")
        with col3:
            otulina_boczna = st.number_input("Otulina boczna [mm]", value=5.0, key="z_ot_boczna")

        col4, col5 = st.columns(2)
        with col4:
            odleglosc_pozioma = st.number_input(
                "Odstęp poziomy między prętami [mm] (clear)",
                value=5.0,
                key="z_odst_poziomy",
            )
        with col5:
            odleglosc_pionowa = st.number_input(
                "Odstęp pionowy między warstwami [mm] (clear)",
                value=5.0,
                key="z_odst_pionowy",
            )

        # Geometria półki w mm
        b_polki_mm = b_polki * 10
        h_polki_mm = h_polki * 10

        # Parametry „własne” (bez średnicy)
        n_wlasne = int(
            st.number_input(
                "Liczba prętów w 1 warstwie [szt.]",
                min_value=0,
                value=4,
                key="z_n_wlasne",
            )
        )
        warstwy_wlasne = int(
            st.number_input(
                "Liczba warstw [szt.]",
                min_value=0,
                value=1,
                key="z_warstwy_wlasne",
            )
        )

        szer_dostepna = b_polki_mm - 2 * otulina_boczna
        wysokosc_dostepna = h_polki_mm - otulina_dolna - otulina_gorna

        # Zapotrzebowanie miejsca (clear spacings)
        req_width = phi_f_mm * n_wlasne + odleglosc_pozioma * (n_wlasne - 1)
        req_height = phi_f_mm * warstwy_wlasne + odleglosc_pionowa * (warstwy_wlasne - 1)

        violations = []
        if szer_dostepna <= 0:
            violations.append("❌ Brak światła na szerokości (otuliny boczne zjadają całą półkę).")
        if wysokosc_dostepna <= 0:
            violations.append("❌ Brak światła na wysokości (otulina górna/dolna zjada całą półkę).")
        if req_width > max(0.0, szer_dostepna):
            violations.append("❌ Pręty **nie mieszczą się na szerokość** dla zadanych otulin i odstępów.")
        if req_height > max(0.0, wysokosc_dostepna):
            violations.append("❌ Pręty/warstwy **nie mieszczą się na wysokość** dla zadanych otulin i odstępów.")

        fits = (
            szer_dostepna > 0
            and wysokosc_dostepna > 0
            and req_width <= szer_dostepna
            and req_height <= wysokosc_dostepna
        )



        for v in violations:
            if v.startswith("❌"):
                st.error(v)
            else:
                st.warning(v)

        # Pole zbrojenia + obwód (używa phi_f_mm)
        A_pręt = math.pi * (phi_f_mm / 2.0) ** 2 / 100.0  # cm²
        As = A_pręt * n_wlasne * warstwy_wlasne
        O_calk = math.pi * phi_f_mm * n_wlasne * warstwy_wlasne
        liczba_pretow = n_wlasne * warstwy_wlasne

        st.session_state["flange_summary"] = {
            "As_cm2": float(As),  # cm²
            "n_bars": int(liczba_pretow),
            "perimeter_mm": float(O_calk),  # mm
        }


    with col_right:
        fig_rebar = go.Figure()
        fig_rebar.update_layout(
            plot_bgcolor="black",
            paper_bgcolor="black",
            font=dict(color="white"),
            width=None,
            height=500,
            showlegend=False,
        )

        # prostokąt półki
        fig_rebar.add_shape(
            type="rect",
            x0=0,
            y0=0,
            x1=b_polki_mm,
            y1=h_polki_mm,
            line=dict(color="white", width=2),
            fillcolor="rgba(80,80,80,0.4)",
        )

        fi = phi_f_mm
        n_pretow = n_wlasne
        warstwy = warstwy_wlasne
        fits_flag = fits

        if not fits_flag:
            st.warning(
                "⚠️ Zbrojenie nie mieści się geometrycznie przy zadanych otulinach/odstępach — "
                "rysunek może nie odzwierciedlać rzeczywistego układu."
            )

        if n_pretow > 0 and warstwy > 0:
            x0 = otulina_boczna + fi / 2.0
            y0 = otulina_dolna + fi / 2.0
            szer_dostepna_vis = max(0.0, b_polki_mm - 2 * otulina_boczna)
            wysokosc_dostepna_vis = max(0.0, h_polki_mm - otulina_dolna - otulina_gorna)

            if n_pretow == 1:
                x_poz = [b_polki_mm / 2.0]
            else:
                x_poz = [
                    x0 + i * (szer_dostepna_vis - fi) / (n_pretow - 1)
                    for i in range(n_pretow)
                ]

            y0 = otulina_dolna + fi / 2.0
            krok_pion = fi + odleglosc_pionowa
            limit_top_center = h_polki_mm - otulina_gorna - fi / 2.0

            if krok_pion > 0 and (limit_top_center - y0) >= 0:
                max_layers_vis = 1 + int((limit_top_center - y0) // krok_pion)
            else:
                max_layers_vis = 0

            layers_to_draw = max(0, min(warstwy, max_layers_vis))

            if layers_to_draw <= 0:
                y_poz = []
            else:
                y_poz = [y0 + j * krok_pion for j in range(layers_to_draw)]

            for y in y_poz:
                for x in x_poz:
                    fig_rebar.add_shape(
                        type="circle",
                        x0=x - fi / 2.0,
                        y0=y - fi / 2.0,
                        x1=x + fi / 2.0,
                        y1=y + fi / 2.0,
                        line=dict(color="red", width=1.5),
                        fillcolor="red",
                    )

        fig_rebar.update_xaxes(
            title="Szerokość półki [mm]",
            range=[-10, b_polki_mm + 10],
            scaleanchor="y",
            scaleratio=1,
        )
        fig_rebar.update_yaxes(
            title="Wysokość półki [mm]",
            range=[-10, h_polki_mm + 10],
        )
        st.plotly_chart(fig_rebar, use_container_width=True)

    # === 3) Górne zbrojenie w środniku – własna definicja (bez pola średnicy) ===
    st.subheader("Rozmieszczenie prętów w środniku")

    if h_srodnika <= 0 or b_srodnika <= 0:
        st.warning("⚠️ Szerokość lub wysokość środnika wynosi 0 – nie można zdefiniować zbrojenia w środniku.")
    else:
        col_ws_left, col_ws_right = st.columns([1, 1])

        with col_ws_left:
            colw1, colw2, colw3 = st.columns(3)
            with colw1:
                otulina_gorna_s = st.number_input(
                    "Otulina górna [mm] (od góry środnika)",
                    value=5.0,
                    key="s_ot_gorna"
                )
            with colw2:
                otulina_dolna_s = st.number_input(
                    "Otulina dolna [mm] (od dołu środnika)",
                    value=5.0,
                    key="s_ot_dolna"
                )
            with colw3:
                otulina_boczna_s = st.number_input(
                    "Otulina boczna [mm] (dla środnika)",
                    value=5.0,
                    key="s_ot_boczna"
                )

            colw4, colw5 = st.columns(2)
            with colw4:
                odleglosc_pozioma_s = st.number_input(
                    "Odstęp poziomy między prętami [mm] (clear, w środniku)",
                    value=5.0,
                    key="s_odst_poziomy",
                )
            with colw5:
                odleglosc_pionowa_s = st.number_input(
                    "Odstęp pionowy między warstwami [mm] (clear, w środniku)",
                    value=5.0,
                    key="s_odst_pionowy",
                )

            # Geometria środnika w mm
            b_srodnika_mm = b_srodnika * 10.0
            h_srodnika_mm = h_srodnika * 10.0

            n_wlasne_s = int(
                st.number_input(
                    "Liczba prętów w 1 warstwie (środnik) [szt.]",
                    min_value=0,
                    value=2,
                    key="s_n_wlasne",
                )
            )
            warstwy_wlasne_s = int(
                st.number_input(
                    "Liczba warstw (środnik) [szt.]",
                    min_value=0,
                    value=1,
                    key="s_warstwy_wlasne",
                )
            )

            szer_dostepna_s = b_srodnika_mm - 2 * otulina_boczna_s
            wysokosc_dostepna_s = h_srodnika_mm - otulina_gorna_s - otulina_dolna_s

            # Zapotrzebowanie miejsca (clear spacings)
            req_width_s = phi_w_mm * n_wlasne_s + odleglosc_pozioma_s * max(0, n_wlasne_s - 1)
            req_height_s = phi_w_mm * warstwy_wlasne_s + odleglosc_pionowa_s * max(0, warstwy_wlasne_s - 1)

            violations_s = []
            if szer_dostepna_s <= 0:
                violations_s.append("❌ Brak światła na szerokości w środniku (otuliny boczne zjadają cały przekrój).")
            if wysokosc_dostepna_s <= 0:
                violations_s.append(
                    "❌ Brak światła na wysokości w środniku (otuliny górna/dolna zjadają cały przekrój).")
            if n_wlasne_s > 0 and warstwy_wlasne_s > 0:
                if req_width_s > max(0.0, szer_dostepna_s):
                    violations_s.append(
                        "❌ Pręty **nie mieszczą się na szerokość** w środniku przy zadanych otulinach/odstępach.")
                if req_height_s > max(0.0, wysokosc_dostepna_s):
                    violations_s.append(
                        "❌ Pręty/warstwy **nie mieszczą się na wysokość** w środniku przy zadanych otulinach/odstępach.")

            fits_s = (
                    n_wlasne_s > 0
                    and warstwy_wlasne_s > 0
                    and szer_dostepna_s > 0
                    and wysokosc_dostepna_s > 0
                    and req_width_s <= max(0.0, szer_dostepna_s)
                    and req_height_s <= max(0.0, wysokosc_dostepna_s)
            )

            for v in violations_s:
                if v.startswith("❌"):
                    st.error(v)
                else:
                    st.warning(v)

            # Pole zbrojenia górnego w środniku + obwód – używa phi_w_mm
            A_pręt_s = math.pi * (phi_w_mm / 2.0) ** 2 / 100.0  # cm²
            As_s = A_pręt_s * n_wlasne_s * warstwy_wlasne_s
            O_calk_s = math.pi * phi_w_mm * n_wlasne_s * warstwy_wlasne_s
            liczba_pretow_s = n_wlasne_s * warstwy_wlasne_s



        with col_ws_right:
            fig_rebar_s = go.Figure()
            fig_rebar_s.update_layout(
                plot_bgcolor="black",
                paper_bgcolor="black",
                font=dict(color="white"),
                width=None,
                height=500,
                showlegend=False,
            )

            # prostokąt środnika (lokalny układ: 0 = dół środnika, h_srodnika_mm = góra środnika)
            fig_rebar_s.add_shape(
                type="rect",
                x0=0,
                y0=0,
                x1=b_srodnika_mm,
                y1=h_srodnika_mm,
                line=dict(color="white", width=2),
                fillcolor="rgba(80,80,80,0.4)",
            )

            fi_s = phi_w_mm
            n_pretow_s = n_wlasne_s
            warstwy_s = warstwy_wlasne_s

            if not fits_s and n_pretow_s > 0 and warstwy_s > 0:
                st.warning(
                    "⚠️ Zbrojenie górne w środniku nie mieści się geometrycznie przy zadanych otulinach/odstępach — "
                    "nie jest rysowane na przekroju."
                )

            # Rysujemy pręty TYLKO jeśli się mieszczą i liczby > 0
            if fits_s and n_pretow_s > 0 and warstwy_s > 0:
                szer_dostepna_vis_s = max(0.0, b_srodnika_mm - 2 * otulina_boczna_s)

                # Poziomo: rozkład po szerokości środnika
                if n_pretow_s == 1:
                    x_poz_s = [b_srodnika_mm / 2.0]
                else:
                    x0_s = otulina_boczna_s + fi_s / 2.0
                    x_poz_s = [
                        x0_s + i * (szer_dostepna_vis_s - fi_s) / (n_pretow_s - 1)
                        for i in range(n_pretow_s)
                    ]

                # Pionowo: warstwy LICZONE OD GÓRY w dół (dociągnięte do górnego lica)
                y_top_center = h_srodnika_mm - otulina_gorna_s - fi_s / 2.0
                krok_pion_s = fi_s + odleglosc_pionowa_s

                y_poz_s = [
                    y_top_center - j * krok_pion_s
                    for j in range(warstwy_s)
                ]

                # Filtrujemy, żeby nie wyjść poniżej otulina_dolna_s
                min_center = otulina_dolna_s + fi_s / 2.0
                y_poz_s = [y for y in y_poz_s if y >= min_center]

                for y in y_poz_s:
                    for x in x_poz_s:
                        fig_rebar_s.add_shape(
                            type="circle",
                            x0=x - fi_s / 2.0,
                            y0=y - fi_s / 2.0,
                            x1=x + fi_s / 2.0,
                            y1=y + fi_s / 2.0,
                            line=dict(color="red", width=1.5),
                            fillcolor="red",
                        )

            fig_rebar_s.update_xaxes(
                title="Szerokość środnika [mm]",
                range=[-10, b_srodnika_mm + 10],
                scaleanchor="y",
                scaleratio=1,
            )
            fig_rebar_s.update_yaxes(
                title="Wysokość środnika [mm] (0 = dół)",
                range=[-10, h_srodnika_mm + 10],
            )
            st.plotly_chart(fig_rebar_s, use_container_width=True)


# =====================================================
# PODSUMOWANIE ZBROJENIA
# =====================================================

st.subheader("Podsumowanie zbrojenia")

# --- dane z półki (dolne zbrojenie) ---
# As, liczba_pretow, O_calk powinny być policzone wcześniej
try:
    As_f = float(As)                 # cm²
    n_f = int(liczba_pretow)         # szt.
    O_f = float(O_calk)              # mm
except Exception:
    As_f, n_f, O_f = 0.0, 0, 0.0

# --- dane ze środnika (górne zbrojenie) ---
# As_s, liczba_pretow_s, O_calk_s też są policzone w sekcji środnika
try:
    As_w = float(As_s)               # cm²
    n_w = int(liczba_pretow_s)       # szt.
    O_w = float(O_calk_s)            # mm
except Exception:
    As_w, n_w, O_w = 0.0, 0, 0.0

# --- łącznie ---
As_total = As_f + As_w
n_total = n_f + n_w
O_total = O_f + O_w

df_rebar_summary = pd.DataFrame(
    [
        ["Półka",    f"{As_f:.2f}",    n_f,     f"{O_f:.0f}"],
        ["Środnik",  f"{As_w:.2f}",    n_w,     f"{O_w:.0f}"],
        ["Łącznie",  f"{As_total:.2f}", n_total, f"{O_total:.0f}"],
    ],
    columns=["Element", "As [cm²]", "Liczba prętów", "Obwód [mm]"],
)

# bez kolumny z indeksami 0/1/2 – jako nagłówek bierzemy "Element"
st.table(df_rebar_summary.set_index("Element"))

# ============================================================
# SEKCJA – WYTRZYMAŁOŚĆ NA ŚCINANIE
# ============================================================

st.markdown("---")
st.header("Wytrzymałość na ścinanie")

# ---- pomocnicze funkcje dla procedur ----

def shear_capacity_ACI(bw_mm, d_mm, Af_mm2, Ef_GPa, Ec_GPa, fck_MPa) -> float:
    """Nośność na ścinanie wg ACI 440 (bez współczynników bezpieczeństwa). Zwraca Pmax [kN]."""
    try:
        if bw_mm <= 0 or d_mm <= 0 or fck_MPa <= 0 or Af_mm2 <= 0 or Ef_GPa <= 0 or Ec_GPa <= 0:
            return float("nan")

        rho_f = Af_mm2 / (bw_mm * d_mm)      # udział zbrojenia FRP
        eta_f = Ef_GPa / Ec_GPa              # Ef/Ec

        k = math.sqrt(max(0.0, 2.0 * eta_f * rho_f + (eta_f * rho_f) ** 2))
        Vc_N = (2.0 / 5.0) * math.sqrt(fck_MPa) * bw_mm * d_mm * k
        Pmax_kN = 2.0 * Vc_N / 1000.0        # dwa obciążenia P/2
        return Pmax_kN
    except Exception:
        return float("nan")


def shear_capacity_JSCE(bw_mm, d_mm, Af_mm2, Ef_GPa, fck_MPa) -> float:
    """Nośność na ścinanie wg JSCE (bez współczynników bezpieczeństwa). Zwraca Pmax [kN]."""
    try:
        if bw_mm <= 0 or d_mm <= 0 or fck_MPa <= 0 or Af_mm2 <= 0 or Ef_GPa <= 0:
            return float("nan")

        rho_f = Af_mm2 / (bw_mm * d_mm)
        E0_GPa = 200.0

        beta_d = min(1.5, (1000.0 / max(1.0, d_mm)) ** 0.25)
        x = 100.0 * rho_f * (Ef_GPa / E0_GPa)
        beta_p = min(1.5, x ** (1.0 / 3.0)) if x > 0 else 0.0

        tau_bd_MPa = min(0.2 * math.sqrt(fck_MPa), 0.72)

        Vc_N = beta_d * beta_p * tau_bd_MPa * bw_mm * d_mm
        Pmax_kN = 2.0 * Vc_N / 1000.0
        return Pmax_kN
    except Exception:
        return float("nan")


def shear_capacity_CSA(bw_mm, h_mm, d_mm, Af_mm2, Ef_GPa, fck_MPa) -> float:
    """
    Nośność na ścinanie wg CSA (wzory 1.21–1.23, 1.24, 1.25) bez współczynników bezpieczeństwa.
    Zwraca Pmax [kN] dla schematu z dwoma siłami P/2.
    """
    try:
        if bw_mm <= 0 or d_mm <= 0 or h_mm <= 0 or fck_MPa <= 0 or Af_mm2 <= 0 or Ef_GPa <= 0:
            return float("nan")

        lambda_c = 1.0   # beton zwykły
        phi_c = 1.0      # brak współczynników bezpieczeństwa

        # dv = min(0.9d ; 0.72h)
        dv_mm = min(0.9 * d_mm, 0.72 * h_mm)

        # rho_l = A_l / (bw * d)
        rho_l = Af_mm2 / (bw_mm * d_mm)

        # kr = 1 + (Ef * rho_l)^(1/3)
        kr = 1.0 + (Ef_GPa * rho_l) ** (1.0 / 3.0)

        # km = sqrt(d/dv) <= 1.0  (dla Twojego schematu MEd = VEd*dv)
        km = math.sqrt(d_mm / dv_mm)
        km = min(km, 1.0)

        # "surowe" Vc z (1.21)
        Vc_raw_N = (
            0.05
            * lambda_c
            * phi_c
            * km
            * kr
            * (fck_MPa ** (1.0 / 3.0))
            * bw_mm
            * dv_mm
        )

        # ograniczenia z (1.23): 0.11*sqrt(fck)*bw*dv <= Vc <= 0.22*sqrt(fck)*bw*dv
        Vc_min_N = 0.11 * phi_c * math.sqrt(fck_MPa) * bw_mm * dv_mm
        Vc_max_N = 0.22 * phi_c * math.sqrt(fck_MPa) * bw_mm * dv_mm

        Vc_N = min(max(Vc_raw_N, Vc_min_N), Vc_max_N)

        Pmax_kN = 2.0 * Vc_N / 1000.0
        return Pmax_kN
    except Exception:
        return float("nan")


# ---- właściwe obliczenia sekcji ----
try:
    # Dane geometryczne przekroju (T-belka) w cm
    bw_cm = b_srodnika
    H_cm = h_polki + h_srodnika

    # Dane zbrojenia z bazy (pręt w półce – rozciągany)
    rebar_info = st.session_state.get("rebar_bar", {}) or {}
    phi_mm = float(rebar_info.get("phi_mm", phi_mm))  # jeśli phi_mm jest w zasięgu, użyjemy go

    Ef_GPa = float(rebar_info.get("E_GPa", float("nan")))
    if math.isnan(Ef_GPa) or Ef_GPa <= 0:
        Ef_GPa = 50.0  # sensowna wartość domyślna

    # Beton
    fck_MPa = float(beton_info.get("f_ck", float("nan")))
    Ec_GPa = float(beton_info.get("E_c_GPa", float("nan")))

    # Domyślne wartości, jeśli coś jest NaN
    if math.isnan(fck_MPa) or fck_MPa <= 0:
        fck_MPa = 30.0
    if math.isnan(Ec_GPa) or Ec_GPa <= 0:
        Ec_GPa = 30.0

    # Ilość prętów rozciąganych w półce
    n_w = int(st.session_state.get("z_n_wlasne", 1))
    n_layers = int(st.session_state.get("z_warstwy_wlasne", 1))
    n_flange = max(1, n_w) * max(1, n_layers)

    # Otulina i odstęp pionowy (mm)
    otulina_dolna_mm = float(st.session_state.get("z_ot_dolna", 5.0))
    odst_pion_mm = float(st.session_state.get("z_odst_pionowy", 5.0))

    # Geometria w mm
    bw_mm = bw_cm * 10.0
    H_mm = H_cm * 10.0

    # Położenie środka zbrojenia (średnio po warstwach)
    y1 = otulina_dolna_mm + phi_mm / 2.0
    y2 = y1 + (n_layers - 1) * (phi_mm + odst_pion_mm)
    y_cent = 0.5 * (y1 + y2)

    d_mm = H_mm - y_cent  # efektywna wysokość

    # Pole zbrojenia rozciąganego
    Af_mm2 = n_flange * (math.pi * (phi_mm ** 2) / 4.0)

    # --- ACI ---
    P_ACI_kN = shear_capacity_ACI(bw_mm, d_mm, Af_mm2, Ef_GPa, Ec_GPa, fck_MPa)

    # --- JSCE ---
    P_JSCE_kN = shear_capacity_JSCE(bw_mm, d_mm, Af_mm2, Ef_GPa, fck_MPa)

    # --- CSA ---
    P_CSA_kN = shear_capacity_CSA(bw_mm, H_mm, d_mm, Af_mm2, Ef_GPa, fck_MPa)

    # Tabela z wynikami
    df_shear = pd.DataFrame(
        [
            {"Procedura": "ACI 440", "P,max [kN]": P_ACI_kN},
            {"Procedura": "JSCE",    "P,max [kN]": P_JSCE_kN},
            {"Procedura": "CSA",     "P,max [kN]": P_CSA_kN},
        ]
    )


    st.dataframe(
        df_shear.style.format({"P,max [kN]": "{:.2f}"}),
        hide_index=True,
        use_container_width=True,
    )

    # Domyślna wartość – najbardziej "bezpieczna"
    P_list = [x for x in [P_ACI_kN, P_JSCE_kN, P_CSA_kN] if isinstance(x, (int, float)) and not math.isnan(x)]
    P_default = min(P_list) if P_list else 0.0

    col_choice, col_custom = st.columns([2, 1])

    with col_choice:
        shear_choice = st.radio(
            "Wybór wartości do obliczeń punktacji",
            ["min(ACI, JSCE, CSA)", "ACI 440", "JSCE", "CSA", "Własna wartość"],
            index=0,
            horizontal=True,
            key="shear_choice_mode",
        )

    with col_custom:
        P_custom_kN = st.number_input(
            "P,max własna [kN]",
            min_value=0.0,
            value=float(P_default),
            step=0.5,
            key="shear_P_custom_kN",
        )

    if shear_choice == "min(ACI, JSCE, CSA)":
        P_used_kN = P_default
    elif shear_choice == "ACI 440":
        P_used_kN = P_ACI_kN
    elif shear_choice == "JSCE":
        P_used_kN = P_JSCE_kN
    elif shear_choice == "CSA":
        P_used_kN = P_CSA_kN
    else:
        P_used_kN = P_custom_kN

    # zapis do session_state dla sekcji „Punktacja”
    st.session_state["shear_P_ACI_kN"] = float(P_ACI_kN)
    st.session_state["shear_P_JSCE_kN"] = float(P_JSCE_kN)
    st.session_state["shear_P_CSA_kN"] = float(P_CSA_kN)
    st.session_state["shear_P_used_kN"] = float(P_used_kN)


except Exception as e:
    st.error(f"Nie udało się policzyć nośności na ścinanie: {e}")




# ==============================================================
# FUNKCJA: koszt 1 m³ mieszanki (tak samo jak w "Tabela mieszanek")
# ==============================================================

@st.cache_data(show_spinner=False)
@st.cache_data(show_spinner=False)
def compute_mix_price_usd_per_m3(
    spreadsheet_id: str,
    sheet_recipes: str,
    sheet_materials: str,
    recipe_name: str
) -> tuple[float, pd.DataFrame]:
    """
    Koszt 1 m³ mieszanki:
      - z arkusza 'receptury' bierzemy masa_kgm3 dla składników danej receptury,
      - z arkusza 'materiały' bierzemy:
          * cena_pln  – cena w USD (po prostu wartość z arkusza),
          * cena_za   – 'kg' lub 'l',
          * gestosc_gcm3 – do przeliczenia kg ↔ l,
      - jeśli cena_za zawiera 'kg'  -> ilość_jedn = masa_kgm3,
      - jeśli cena_za zawiera 'l'   -> ilość_jedn = (masa_kgm3 / (rho_kgm3)) * 1000,
      - inaczej traktujemy jak 'kg'.

    Zwraca:
      (cena_m3, df_skladniki), gdzie df_skladniki ma kolumny:
      material, cena_za, cena_jedn, ilosc_jedn, koszt
    """

    if not GS_RECIPES_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets.")

    # --- lokalne pomocnicze parsowanie ---
    def _to_num(x):
        s = str(x).strip()
        if not s:
            return float("nan")
        s = s.replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return float("nan")

    def _strip_apos(x):
        s = str(x)
        return s[1:] if s.startswith("'") else s

    # ---------- RECEPTURY ----------
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)

    ws_rec = ss.worksheet(sheet_recipes)
    values = ws_rec.get_all_values()
    if not values:
        return 0.0, pd.DataFrame(columns=["material", "cena_za", "cena_jedn", "ilosc_jedn", "koszt"])

    header = values[0]
    df_rec = pd.DataFrame(values[1:], columns=header)

    for col in ["recipe_name", "nazwa", "material_id", "masa_kgm3"]:
        if col not in df_rec.columns:
            df_rec[col] = ""

    # tylko nasza receptura, bez SUMMARY/__SUMMARY__
    df_rec = df_rec[df_rec["recipe_name"].astype(str) == str(recipe_name)]
    df_rec = df_rec[~df_rec["nazwa"].astype(str).str.upper().isin(["SUMMARY", "__SUMMARY__"])]

    if df_rec.empty:
        return 0.0, pd.DataFrame(columns=["material", "cena_za", "cena_jedn", "ilosc_jedn", "koszt"])

    df_rec["material_id"] = df_rec["material_id"].apply(_strip_apos)
    df_rec["material_id"] = pd.to_numeric(df_rec["material_id"], errors="coerce")
    df_rec["masa_kgm3"] = df_rec["masa_kgm3"].apply(_to_num).fillna(0.0)

    # ---------- MATERIAŁY ----------
    ws_mat = ss.worksheet(sheet_materials)
    rows_mat = ws_mat.get_all_records(numericise_ignore=["all"])
    df_mat = pd.DataFrame(rows_mat)

    for c in ["id", "nazwa", "gestosc_gcm3", "cena_pln", "cena_za"]:
        if c not in df_mat.columns:
            df_mat[c] = None

    df_mat["id"] = df_mat["id"].apply(_strip_apos)
    df_mat["id"] = pd.to_numeric(df_mat["id"], errors="coerce")
    df_mat["gestosc_gcm3"] = df_mat["gestosc_gcm3"].apply(_to_num)
    df_mat["cena_pln"] = df_mat["cena_pln"].apply(_to_num).fillna(0.0)

    df_mat_id = df_mat.rename(columns={
        "id": "material_id",
        "nazwa": "nazwa_mat",
        "gestosc_gcm3": "rho_gcm3",
        "cena_za": "cena_za_mat",
    })[["material_id", "nazwa_mat", "rho_gcm3", "cena_pln", "cena_za_mat"]]

    # ---------- POŁĄCZENIE ----------
    df_cost = df_rec.merge(df_mat_id, on="material_id", how="left")

    df_cost["masa_kgm3"] = df_cost["masa_kgm3"].fillna(0.0)
    df_cost["rho_gcm3"] = df_cost["rho_gcm3"].apply(_to_num)
    df_cost["rho_kgm3"] = df_cost["rho_gcm3"] * 1000.0

    # ---------- ILOŚĆ JEDNOSTEK W 1 m³ (kg lub l) ----------
    def _calc_amount(row):
        masa = float(row["masa_kgm3"])          # kg/m³
        rho  = float(row["rho_kgm3"]) if row["rho_kgm3"] == row["rho_kgm3"] else float("nan")
        unit_raw = str(row["cena_za_mat"] or "").strip().lower()
        unit = unit_raw.replace("/", "").replace(" ", "")

        # domyślnie jak kg
        if "l" in unit and "kg" not in unit:
            # cena za litr
            if rho is not None and rho == rho and rho > 0:
                vol_m3 = masa / rho           # m³/m³
                return vol_m3 * 1000.0        # l/m³
            else:
                # brak gęstości -> traktujemy jak kg
                return masa
        else:
            # cena za kg (albo brak jednostki)
            return masa

    df_cost["ilosc_jedn"] = df_cost.apply(_calc_amount, axis=1)

    # ---------- KOSZT SKŁADNIKA I KOSZT MIESZANKI ----------
    df_cost["koszt_m3"] = df_cost["ilosc_jedn"] * df_cost["cena_pln"]
    cena_m3 = float(df_cost["koszt_m3"].sum()) if not df_cost.empty else 0.0

    # ---------- TABELKA DO EXPANDERA ----------
    df_break = pd.DataFrame({
        "material": df_cost["nazwa_mat"].fillna(df_cost["nazwa"]).fillna("(?)"),
        "cena_za": df_cost["cena_za_mat"].fillna("kg"),
        "cena_jedn": df_cost["cena_pln"],
        "ilosc_jedn": df_cost["ilosc_jedn"],
        "koszt": df_cost["koszt_m3"],
    })

    return cena_m3, df_break

@st.cache_data(show_spinner=False)
def compute_material_adjustment_pct(
    spreadsheet_id: str,
    sheet_recipes: str,
    sheet_materials: str,
    recipe_name: str,
) -> float:
    """
    Liczy łączną korektę materiałową [%] wg progów schodkowych,
    na podstawie mas składników (masa_kgm3) i kolumny 'atrybut' w arkuszu materiałów.

    Atrybuty (dokładne stringi):
      - "Cement korekta"
      - "Popiół lotny dodatek"
      - "Żużel wiel. dodatek"
      - "Pył krzem. dodatek"
      - pozostałe / puste -> ignorowane dla cementitious
    """

    if not GS_RECIPES_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets.")

    def _to_num(x):
        s = str(x).strip()
        if not s:
            return 0.0
        s = s.replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    def _strip_apos(x):
        s = str(x)
        return s[1:] if s.startswith("'") else s

    # --- pobranie danych z GS ---
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)

    # Receptury
    ws_rec = ss.worksheet(sheet_recipes)
    values = ws_rec.get_all_values()
    if not values:
        return 0.0

    header = values[0]
    df_rec = pd.DataFrame(values[1:], columns=header)

    for col in ["recipe_name", "nazwa", "material_id", "masa_kgm3"]:
        if col not in df_rec.columns:
            df_rec[col] = ""

    df_rec = df_rec[df_rec["recipe_name"].astype(str) == str(recipe_name)]
    df_rec = df_rec[~df_rec["nazwa"].astype(str).str.upper().isin(["SUMMARY", "__SUMMARY__"])]

    if df_rec.empty:
        return 0.0

    df_rec["material_id"] = df_rec["material_id"].apply(_strip_apos)
    df_rec["material_id"] = pd.to_numeric(df_rec["material_id"], errors="coerce")
    df_rec["masa_kgm3"] = df_rec["masa_kgm3"].apply(_to_num).fillna(0.0)

    # Materiały
    ws_mat = ss.worksheet(sheet_materials)
    rows_mat = ws_mat.get_all_records(numericise_ignore=["all"])
    df_mat = pd.DataFrame(rows_mat)

    for c in ["id", "atrybut"]:
        if c not in df_mat.columns:
            df_mat[c] = None

    df_mat["id"] = df_mat["id"].apply(_strip_apos)
    df_mat["id"] = pd.to_numeric(df_mat["id"], errors="coerce")

    df_mat_small = df_mat.rename(columns={"id": "material_id"})[["material_id", "atrybut"]]

    # Join
    df = df_rec.merge(df_mat_small, on="material_id", how="left")

    # --- masy ---
    total_mass = float(df["masa_kgm3"].sum())

    s_attr = df["atrybut"].fillna("").astype(str).str.strip()

    def _mass_for(attr: str) -> float:
        return float(df.loc[s_attr == attr, "masa_kgm3"].sum())

    cement_mass = _mass_for("Cement korekta")
    flyash_mass = _mass_for("Popiół lotny dodatek")
    slag_mass   = _mass_for("Żużel wiel. dodatek")
    silica_mass = _mass_for("Pył krzem. dodatek")

    cementitious_sum = cement_mass + flyash_mass + slag_mass + silica_mass

    # --- progi schodkowe (sumują się) ---
    bonus = 0.0

    # Cement vs TOTAL (all concrete materials)
    if total_mass > 0:
        cement_ratio = cement_mass / total_mass
        if cement_ratio < 0.15:
            bonus += 1.0
        if cement_ratio < 0.10:
            bonus += 1.0
        if cement_ratio < 0.05:
            bonus += 1.0

    # SCM vs cementitious
    if cementitious_sum > 0:
        fly_ratio = flyash_mass / cementitious_sum
        if fly_ratio > 0.20:
            bonus += 1.0
        if fly_ratio > 0.30:
            bonus += 1.0
        if fly_ratio > 0.40:
            bonus += 1.0

        slag_ratio = slag_mass / cementitious_sum
        if slag_ratio > 0.20:
            bonus += 1.0
        if slag_ratio > 0.35:
            bonus += 1.0
        if slag_ratio > 0.50:
            bonus += 1.0

        silica_ratio = silica_mass / cementitious_sum
        if silica_ratio > 0.05:
            bonus += 1.0
        if silica_ratio > 0.10:
            bonus += 1.0

    return float(bonus)


# ============================================================
# SEKCJA – PUNKTACJA
# ============================================================

st.markdown("---")
st.header("Punktacja")

try:
    # ------------------------------------------------------
    # 1) OBJĘTOŚĆ BELKI (bez zbrojenia)
    # ------------------------------------------------------
    A_m2 = (b_polki / 100.0 * h_polki / 100.0) + (b_srodnika / 100.0 * h_srodnika / 100.0)
    L_m = float(L_beam)
    V_beam_m3 = A_m2 * L_m
    V_beam_L = V_beam_m3 * 1000.0

    # ------------------------------------------------------
    # 2) CENA MIESZANKI – z arkuszy (materiały + receptury)
    # ------------------------------------------------------
    sel_recipe_name = None
    if beton_mode == "Wybór z bazy danych":
        sel_recipe_name = st.session_state.get("beton_recipe_name")

    if not GS_RECIPES_READY or not sel_recipe_name:
        st.error("Brak dostępu do Google Sheets lub nie wybrano receptury — nie mogę policzyć ceny mieszanki.")
        mix_price_usd_m3 = 0.0
        df_mix_break = pd.DataFrame()
    else:
        try:
            mix_price_usd_m3, df_mix_break = compute_mix_price_usd_per_m3(
                SPREADSHEET_ID, SHEET_RECIPES, SHEET_MATERIALS, sel_recipe_name
            )
        except Exception as e:
            st.error(f"Nie udało się policzyć ceny mieszanki z arkuszy: {e}")
            mix_price_usd_m3 = 0.0
            df_mix_break = pd.DataFrame()

    # ------------------------------------------------------
    # Korekta materiałowa [%] wg atrybutów
    #   działa na korzyść wyniku -> trzymamy jako wartość ujemną (np. -3)
    # ------------------------------------------------------
    material_adj_pct = 0.0
    if GS_RECIPES_READY and sel_recipe_name:
        try:
            material_adj_pct = compute_material_adjustment_pct(
                SPREADSHEET_ID,
                SHEET_RECIPES,
                SHEET_MATERIALS,
                sel_recipe_name,
            )
            if material_adj_pct != material_adj_pct:  # NaN
                material_adj_pct = 0.0
            material_adj_pct = -abs(material_adj_pct)
        except Exception as e:
            st.error(f"Nie udało się policzyć korekty materiałowej: {e}")
            material_adj_pct = 0.0

    # ------------------------------------------------------
    # Korekta geometryczna [%] dla belki teowej (stałe 15%)
    # ------------------------------------------------------
    geom_adj_pct = 15.0

    # ------------------------------------------------------
    # 3) PARAMETRY BETONU
    # ------------------------------------------------------
    rho_conc_kgm3 = float(beton_info.get("rho", 2400.0))

    # ------------------------------------------------------
    # 4) ZBROJENIE – ilość, objętość, koszt
    # ------------------------------------------------------
    rebar_info = st.session_state.get("rebar_bar", {}) or {}

    phi_mm = float(rebar_info.get("phi_mm", float("nan")))
    Ef_GPa = float(rebar_info.get("E_GPa", float("nan")))
    rho_bar_kgm3 = float(rebar_info.get("rho_kgm3", float("nan")))

    # cena jednostkowa z arkusza
    try:
        price_bar = float(rebar_info.get("price_pln", 0.0))
        if math.isnan(price_bar):
            price_bar = 0.0
    except Exception:
        price_bar = 0.0

    unit_raw = (rebar_info.get("price_unit") or "").strip().lower()
    unit_norm = unit_raw.replace("/", "").replace(" ", "")

    # liczba prętów
    n_flange = int(st.session_state.get("z_n_wlasne", 1)) * int(st.session_state.get("z_warstwy_wlasne", 1))
    n_web = int(st.session_state.get("s_n_wlasne", 1)) * int(st.session_state.get("s_warstwy_wlasne", 1))
    n_bars_tot = int(n_flange + n_web)

    # geometria prętów
    area_bar_m2 = math.pi * ((phi_mm / 1000.0) ** 2) / 4.0
    total_length_m = n_bars_tot * L_m
    V_rebar_m3 = area_bar_m2 * total_length_m
    mass_rebar_kg = V_rebar_m3 * (rho_bar_kgm3 if rho_bar_kgm3 == rho_bar_kgm3 else 0.0)

    # koszt zbrojenia
    if "kg" in unit_norm:
        cost_rebar_usd = price_bar * mass_rebar_kg
    elif "mb" in unit_norm or unit_norm == "m":
        cost_rebar_usd = price_bar * n_bars_tot * L_m
    else:
        cost_rebar_usd = price_bar * n_bars_tot * L_m

    # ------------------------------------------------------
    # 5) CENA MIESZANKI DLA BELKI (z odjęciem objętości prętów)
    # ------------------------------------------------------
    V_conc_net_m3 = max(0.0, V_beam_m3 - V_rebar_m3)
    cost_mix_usd = mix_price_usd_m3 * V_conc_net_m3

    # ------------------------------------------------------
    # 6) MASA BELKI
    # ------------------------------------------------------
    mass_conc_kg = V_conc_net_m3 * rho_conc_kgm3
    mass_total_kg = mass_conc_kg + mass_rebar_kg

    # ------------------------------------------------------
    # 7) KOSZT TRANSPORTU – 0.01 USD za każdy rozpoczęty lb
    # ------------------------------------------------------
    lb_per_kg = 2.20462262185
    mass_lb = mass_total_kg * lb_per_kg
    transport_usd = 0.01 * math.ceil(mass_lb)

    # ------------------------------------------------------
    # 8) WYTRZYMAŁOŚĆ (P,max)
    # ------------------------------------------------------
    bw_mm = b_srodnika * 10.0
    H_mm = (h_polki + h_srodnika) * 10.0

    otulina_dolna_mm = float(st.session_state.get("z_ot_dolna", 5.0))
    odst_pion_mm = float(st.session_state.get("z_odst_pionowy", 5.0))

    y1 = otulina_dolna_mm + phi_mm / 2.0
    y2 = y1 + (int(st.session_state.get("z_warstwy_wlasne", 1)) - 1) * (phi_mm + odst_pion_mm)
    y_cent = (y1 + y2) / 2.0
    d_mm = H_mm - y_cent

    Af_mm2 = n_flange * (math.pi * (phi_mm ** 2) / 4.0)
    rho_f = Af_mm2 / (bw_mm * d_mm) if (bw_mm > 0 and d_mm > 0) else 0.0
    eta_f = (Ef_GPa / Ec_GPa) if (Ec_GPa and Ec_GPa > 0) else 0.0

    k_fac = math.sqrt(max(0.0, 2.0 * eta_f * rho_f + (eta_f * rho_f) ** 2))
    Vc_ACI_N = (2.0 / 5.0) * math.sqrt(max(0.0, fck_MPa)) * bw_mm * d_mm * k_fac
    P_ACI_kN_loc = 2.0 * Vc_ACI_N / 1000.0

    E0_GPa = 200.0
    beta_d = min(1.5, (1000.0 / max(1.0, d_mm)) ** 0.25)
    beta_p = min(1.5, (100.0 * rho_f * (Ef_GPa / E0_GPa)) ** (1.0 / 3.0)) if (Ef_GPa and Ef_GPa > 0) else 0.0
    tau_bd_MPa = min(0.2 * math.sqrt(max(0.0, fck_MPa)), 0.72)
    Vc_JSCE_N = beta_d * beta_p * tau_bd_MPa * bw_mm * d_mm
    P_JSCE_kN_loc = 2.0 * Vc_JSCE_N / 1000.0

    lambda_c = 1.0
    k_v = 0.05
    k_f = (Ef_GPa / Ec_GPa) ** (1.0 / 3.0) if Ec_GPa and Ec_GPa > 0 else 0.0
    k_rho = rho_f ** (1.0 / 3.0) if rho_f > 0 else 0.0
    Vc_CSA_N = lambda_c * k_v * k_f * k_rho * (fck_MPa ** (1.0 / 3.0)) * bw_mm * d_mm
    P_CSA_kN_loc = 2.0 * Vc_CSA_N / 1000.0

    P_min_loc = min(P_ACI_kN_loc, P_JSCE_kN_loc, P_CSA_kN_loc)

    P_used_ext = st.session_state.get("shear_P_used_kN", None)
    if isinstance(P_used_ext, (int, float)) and P_used_ext > 0:
        P_used_kN = float(P_used_ext)
    else:
        P_used_kN = P_min_loc

    # ------------------------------------------------------
    # KOSZTY MATERIAŁÓW:
    #  - BRUTTO: bez korekty materiałowej
    #  - NETTO: po korekcie materiałowej (material_adj_pct jest ujemny)
    # ------------------------------------------------------
    cost_materials_brutto_usd = cost_mix_usd + cost_rebar_usd
    cost_materials_netto_usd = cost_materials_brutto_usd * (1.0 + material_adj_pct / 100.0)

    # ------------------------------------------------------
    # Cena belki, brutto [USD] (jak dotychczas)
    # ------------------------------------------------------
    price_beam_brutto_usd = cost_materials_brutto_usd + transport_usd

    # ------------------------------------------------------
    # Cena belki, netto [USD]
    #   = koszt materiałów netto + korekta geometryczna (15% od kosztu materiałów netto) + transport
    # ------------------------------------------------------
    geom_correction_usd = cost_materials_netto_usd * (geom_adj_pct / 100.0)
    price_beam_netto_usd = cost_materials_netto_usd + geom_correction_usd + transport_usd

    # ------------------------------------------------------
    # WYNIK USD/kN (TERAZ z ceny belki NETTO)
    # ------------------------------------------------------
    wynik_usd_per_kN = (price_beam_netto_usd / P_used_kN) if P_used_kN > 0 else float("inf")

    df_score = pd.DataFrame(
        [
            {
                "Łączna obj. belki [l]": V_beam_L,
                "Cena mieszanki / belkę [USD]": cost_mix_usd,
                "Łączna ilość prętów": n_bars_tot,
                "Łączna cena zbrojenia [USD]": cost_rebar_usd,
                "Całkowita masa belki [kg]": mass_total_kg,

                "Koszt materiałów, brutto [USD]": cost_materials_brutto_usd,
                "Korekta materiałowa [%]": material_adj_pct,
                "Koszt materiałów, netto [USD]": cost_materials_netto_usd,

                "Korekta geometryczna [%]": geom_adj_pct,
                "Koszta transportu [USD]": transport_usd,
                "Cena belki, brutto [USD]": price_beam_brutto_usd,
                "Cena belki, netto [USD]": price_beam_netto_usd,

                "P,max [kN]": P_used_kN,
                "Wynik [USD/kN]": wynik_usd_per_kN,
            }
        ]
    )

    st.session_state["score_row"] = df_score.iloc[0].to_dict()

    st.dataframe(
        df_score.style.format(
            {
                "Łączna obj. belki [l]": "{:.1f}",
                "Cena mieszanki / belkę [USD]": "{:.2f}",
                "Łączna cena zbrojenia [USD]": "{:.2f}",
                "Całkowita masa belki [kg]": "{:.2f}",

                "Koszt materiałów, brutto [USD]": "{:.2f}",
                "Korekta materiałowa [%]": "{:.0f}",
                "Koszt materiałów, netto [USD]": "{:.2f}",

                "Korekta geometryczna [%]": "{:.0f}",
                "Koszta transportu [USD]": "{:.2f}",
                "Cena belki, brutto [USD]": "{:.2f}",
                "Cena belki, netto [USD]": "{:.2f}",

                "P,max [kN]": "{:.1f}",
                "Wynik [USD/kN]": "{:.2f}",
            },
            na_rep="—",
        ),
        use_container_width=True,
        hide_index=True,
    )

except Exception as e:
    st.error(f"Nie udało się policzyć punktacji: {e}")


# ============================================================
# ZAPIS BELKI DO GOOGLE SHEETS (arkusz: "belki tpd")
#   - nadpisywanie po nazwie
#   - Geometria = "tpd"
#   - zapis INPUT_* (żeby dało się odtworzyć belkę w edytorze)
#   - nagłówki zawsze wymuszone (_ensure_headers)
# ============================================================

from typing import Any, List

SHEET_BEAMS_TPD = st.secrets.get("SHEET_BEAMS_TPD", "belki tpd")

# ---------------------------
# Nagłówki wyników (jak było)
# ---------------------------
BEAM_HEADERS = [
    "ID",
    "Nazwa belki",
    "Receptura betonu",
    "Geometria",
    "P_ACI_440_kN",
    "P_JSCE_kN",
    "P_CSA_kN",
    "P_custom_kN",
    "P_min_proc_kN",
    "P_worst_all_kN",
    "Wynik_min_proc_USD_per_kN",
    "Wynik_custom_USD_per_kN",
    "Wynik_worst_all_USD_per_kN",
    # --- Pola z tabeli Punktacja ---
    "Łączna obj. belki [l]",
    "Cena mieszanki / belkę [USD]",
    "Łączna ilość prętów",
    "Łączna cena zbrojenia [USD]",
    "Całkowita masa belki [kg]",
    "Koszt materiałów, brutto [USD]",
    "Korekta materiałowa [%]",
    "Koszt materiałów, netto [USD]",
    "Korekta geometryczna [%]",
    "Koszta transportu [USD]",
    "Cena belki, brutto [USD]",
    "Cena belki, netto [USD]",
    "P,max used (worst-case) [kN]",
    "Wynik used (worst-case) [USD/kN]",
]

# ---------------------------
# INPUTS do odtworzenia belki (tpd)
# ---------------------------
INPUT_HEADERS = [
    "INPUT_L_m",

    # geometria T (cm)
    "INPUT_bf_cm",
    "INPUT_hf_cm",
    "INPUT_bw_cm",
    "INPUT_hw_cm",

    # ograniczenia masy
    "INPUT_masa_min_kg",
    "INPUT_masa_max_kg",

    # beton
    "INPUT_beton_mode",          # "gsheet" / "manual"
    "INPUT_beton_recipe_name",
    "INPUT_fck_MPa",
    "INPUT_fctm_MPa",
    "INPUT_rho_kgm3",
    "INPUT_Ec_GPa",

    # pręt GFRP
    "INPUT_gfrp_bar_id",
    "INPUT_phi_mm",
    "INPUT_Ef_GPa",

    # zbrojenie dolne w półce (z_*)
    "INPUT_z_ot_dolna_mm",
    "INPUT_z_ot_gorna_mm",
    "INPUT_z_ot_boczna_mm",
    "INPUT_z_odst_poziomy_mm",
    "INPUT_z_odst_pionowy_mm",
    "INPUT_z_n_wlasne",
    "INPUT_z_warstwy_wlasne",

    # zbrojenie górne w środniku (s_*)
    "INPUT_s_ot_gorna_mm",
    "INPUT_s_ot_dolna_mm",
    "INPUT_s_ot_boczna_mm",
    "INPUT_s_odst_poziomy_mm",
    "INPUT_s_odst_pionowy_mm",
    "INPUT_s_n_wlasne",
    "INPUT_s_warstwy_wlasne",

    # wybór ścinania
    "INPUT_shear_choice_mode",
    "INPUT_shear_P_custom_kN",
]

BEAM_HEADERS = BEAM_HEADERS + INPUT_HEADERS


# ---------------------------
# Helpery arkusza
# ---------------------------
def _normalize_name(s: str) -> str:
    return " ".join(str(s).split()).strip().lower()

def _get_or_create_worksheet(ss: Any, sheet_name: str, headers: List[str]) -> Any:
    try:
        return ss.worksheet(sheet_name)
    except Exception:
        ws = ss.add_worksheet(title=sheet_name, rows=2000, cols=max(20, len(headers)))
        ws.update("A1", [headers], value_input_option="RAW")
        return ws

def _update_sheet_atomic(ws, headers: List[str], df: pd.DataFrame):
    ws.clear()
    ws.update("A1", [headers], value_input_option="RAW")
    data = df.astype(object).values.tolist()
    if not data:
        return

    CHUNK = 800
    start_row = 2
    start_col = 1
    ncols = len(headers)

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

def _ensure_headers(ws, headers: List[str]):
    """
    Gwarantuje nagłówki.
    Jeśli arkusz ma dane bez nagłówków -> przepisuje: headers + stare wiersze jako dane.
    """
    vals = ws.get_all_values() or []
    if not vals:
        ws.update("A1", [headers], value_input_option="RAW")
        return

    if vals[0] == headers:
        return

    width = len(headers)
    fixed = []
    for row in vals:
        r = list(row)
        if len(r) < width:
            r += [""] * (width - len(r))
        else:
            r = r[:width]
        fixed.append(r)

    df_existing = pd.DataFrame(fixed, columns=headers)
    _update_sheet_atomic(ws, headers, df_existing)

def _sheet_to_df(ws, headers: List[str]) -> pd.DataFrame:
    vals = ws.get_all_values() or []
    width = len(headers)
    if not vals:
        return pd.DataFrame(columns=headers)

    fixed = []
    for row in vals:
        r = list(row)
        if len(r) < width:
            r += [""] * (width - len(r))
        else:
            r = r[:width]
        fixed.append(r)

    if fixed and fixed[0] == headers:
        fixed = fixed[1:]

    return pd.DataFrame(fixed, columns=headers)

def _next_id_from_df(existing_df: pd.DataFrame) -> int:
    if existing_df is None or existing_df.empty or "ID" not in existing_df.columns:
        return 1
    ids = []
    for x in existing_df["ID"].tolist():
        try:
            ids.append(int(float(str(x).replace(",", ".").strip())))
        except Exception:
            pass
    return (max(ids) + 1) if ids else 1

def _safe_float(x, default=float("nan")) -> float:
    try:
        return float(x)
    except Exception:
        return default


# ---------------------------
# Budowanie wiersza do zapisu (tpd)
# ---------------------------
def _build_beam_row(beam_name: str) -> List[Any]:
    beton_info = st.session_state.get("beton_dane", {}) or {}
    rebar_info = st.session_state.get("rebar_bar", {}) or {}

    recipe_name = st.session_state.get("beton_recipe_name") or beton_info.get("klasa") or ""

    # ---- ścinanie ----
    P_ACI = _safe_float(st.session_state.get("shear_P_ACI_kN", float("nan")))
    P_JSCE = _safe_float(st.session_state.get("shear_P_JSCE_kN", float("nan")))
    P_CSA = _safe_float(st.session_state.get("shear_P_CSA_kN", float("nan")))
    P_custom = _safe_float(st.session_state.get("shear_P_custom_kN", float("nan")))

    P_list_proc = [p for p in [P_ACI, P_JSCE, P_CSA] if (p == p) and p > 0]
    P_min_proc = min(P_list_proc) if P_list_proc else float("nan")

    candidates = []
    if (P_min_proc == P_min_proc) and P_min_proc > 0:
        candidates.append(P_min_proc)
    if (P_custom == P_custom) and P_custom > 0:
        candidates.append(P_custom)
    if not candidates:
        raise RuntimeError("Brak poprawnej wartości P,max (procedury i/lub własna wartość są puste/niepoprawne).")
    P_worst_all = min(candidates)

    # ---- punktacja ----
    score_row = st.session_state.get("score_row")
    if not score_row:
        raise RuntimeError("Brak danych z sekcji 'Punktacja' (score_row).")

    price_netto = _safe_float(score_row.get("Cena belki, netto [USD]", float("nan")))
    if (price_netto != price_netto) or price_netto <= 0:
        raise RuntimeError("Brak poprawnej wartości 'Cena belki, netto [USD]' w punktacji.")

    wynik_min_proc = (price_netto / P_min_proc) if (P_min_proc == P_min_proc and P_min_proc > 0) else float("nan")
    wynik_custom = (price_netto / P_custom) if (P_custom == P_custom and P_custom > 0) else float("nan")
    wynik_worst = price_netto / P_worst_all if P_worst_all > 0 else float("inf")

    row = [
        "",  # ID
        beam_name.strip(),
        str(recipe_name),
        "tpd",  # ✅ geometria T
        P_ACI, P_JSCE, P_CSA,
        P_custom,
        P_min_proc,
        P_worst_all,
        wynik_min_proc,
        wynik_custom,
        wynik_worst,
        # --- punktacja ---
        score_row.get("Łączna obj. belki [l]", ""),
        score_row.get("Cena mieszanki / belkę [USD]", ""),
        score_row.get("Łączna ilość prętów", ""),
        score_row.get("Łączna cena zbrojenia [USD]", ""),
        score_row.get("Całkowita masa belki [kg]", ""),
        score_row.get("Koszt materiałów, brutto [USD]", ""),
        score_row.get("Korekta materiałowa [%]", ""),
        score_row.get("Koszt materiałów, netto [USD]", ""),
        score_row.get("Korekta geometryczna [%]", ""),
        score_row.get("Koszta transportu [USD]", ""),
        score_row.get("Cena belki, brutto [USD]", ""),
        score_row.get("Cena belki, netto [USD]", ""),
        P_worst_all,
        wynik_worst,
    ]

    # ---- INPUTS (tpd) ----
    beton_mode_src = str(st.session_state.get("beton_mode", "Wybór z bazy danych"))
    beton_recipe_name = st.session_state.get("beton_recipe_name") or beton_info.get("klasa") or ""

    # Geometria T (zakładamy, że w appce masz te zmienne jak wcześniej)
    # L_beam [m], b_polki [cm], h_polki [cm], b_srodnika [cm], h_srodnika [cm]
    row += [
        float(L_beam),
        float(b_polki),
        float(h_polki),
        float(b_srodnika),
        float(h_srodnika),

        float(masa_min),
        float(masa_max),

        beton_mode_src,
        str(beton_recipe_name),
        _safe_float(beton_info.get("f_ck", float("nan"))),
        _safe_float(beton_info.get("f_ctm", float("nan"))),
        _safe_float(beton_info.get("rho", float("nan"))),
        _safe_float(beton_info.get("E_c_GPa", float("nan"))),

        int(rebar_info.get("id")) if rebar_info.get("id") is not None else "",
        _safe_float(rebar_info.get("phi_mm", float("nan"))),
        _safe_float(rebar_info.get("E_GPa", float("nan"))),

        # dolne zbrojenie w półce (z_*)
        _safe_float(st.session_state.get("z_ot_dolna", 5.0)),
        _safe_float(st.session_state.get("z_ot_gorna", 5.0)),
        _safe_float(st.session_state.get("z_ot_boczna", 5.0)),
        _safe_float(st.session_state.get("z_odst_poziomy", 5.0)),
        _safe_float(st.session_state.get("z_odst_pionowy", 5.0)),
        int(st.session_state.get("z_n_wlasne", 0)),
        int(st.session_state.get("z_warstwy_wlasne", 0)),

        # górne zbrojenie w środniku (s_*)
        _safe_float(st.session_state.get("s_ot_gorna", 5.0)),
        _safe_float(st.session_state.get("s_ot_dolna", 5.0)),
        _safe_float(st.session_state.get("s_ot_boczna", 5.0)),
        _safe_float(st.session_state.get("s_odst_poziomy", 5.0)),
        _safe_float(st.session_state.get("s_odst_pionowy", 5.0)),
        int(st.session_state.get("s_n_wlasne", 0)),
        int(st.session_state.get("s_warstwy_wlasne", 0)),

        # wybór ścinania
        str(st.session_state.get("shear_choice_mode", "")),
        _safe_float(st.session_state.get("shear_P_custom_kN", float("nan"))),
    ]

    return row


# ---------------------------
# UI: zapis belki (formularz + checkbox nadpisania)
# ---------------------------
st.markdown("---")
st.header("Zapis belki")

disabled_gs = (not GS_RECIPES_READY) or (not SPREADSHEET_ID)

with st.form(key="save_beam_form", clear_on_submit=False):
    colL, colR = st.columns([1.2, 2])

    with colR:
        beam_name_in = st.text_input("Nazwa belki", key="beam_name_to_save")

    confirm_overwrite_beam = st.checkbox(
        "Nadpisz istniejącą belkę o tej nazwie, jeśli istnieje",
        value=False,
        key="chk_overwrite_beam",
    )

    with colL:
        submit_beam = st.form_submit_button("💾 Zapisz belkę", disabled=disabled_gs)

    if disabled_gs:
        st.info("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account).")

if submit_beam:
    if not beam_name_in.strip():
        st.error("Podaj nazwę belki.")
    else:
        try:
            gc = gspread.authorize(CREDS)
            ss = gc.open_by_key(SPREADSHEET_ID)
            ws_beams = _get_or_create_worksheet(ss, SHEET_BEAMS_TPD, BEAM_HEADERS)

            # ✅ zawsze nagłówki
            _ensure_headers(ws_beams, BEAM_HEADERS)

            existing_df = _sheet_to_df(ws_beams, BEAM_HEADERS)

            if "Nazwa belki" not in existing_df.columns:
                existing_df["Nazwa belki"] = ""

            name_norm = _normalize_name(beam_name_in)
            exists_mask = existing_df["Nazwa belki"].apply(_normalize_name).eq(name_norm)
            exists = bool(exists_mask.any())

            new_row = _build_beam_row(beam_name_in)
            new_row_df = pd.DataFrame([new_row], columns=BEAM_HEADERS)

            if exists and not confirm_overwrite_beam:
                st.warning("Belka o tej nazwie już istnieje. Zaznacz checkbox, aby **nadpisać**, i kliknij ponownie.")
            else:
                if exists:
                    keep_df = existing_df.loc[~exists_mask].copy()
                    new_id = _next_id_from_df(keep_df)
                    new_row_df.loc[0, "ID"] = new_id

                    final_df = pd.concat([keep_df, new_row_df], ignore_index=True)
                    _update_sheet_atomic(ws_beams, BEAM_HEADERS, final_df)

                    st.success(f"Nadpisano belkę „{beam_name_in}” w arkuszu „{SHEET_BEAMS_TPD}”. ID = {new_id}")
                else:
                    new_id = _next_id_from_df(existing_df)
                    new_row_df.loc[0, "ID"] = new_id

                    try:
                        ws_beams.append_rows(new_row_df.values.tolist(), value_input_option="RAW")
                    except Exception:
                        final_df = pd.concat([existing_df, new_row_df], ignore_index=True)
                        _update_sheet_atomic(ws_beams, BEAM_HEADERS, final_df)

                    st.success(f"Zapisano belkę „{beam_name_in}” do arkusza „{SHEET_BEAMS_TPD}”. ID = {new_id}")

        except Exception as e:
            st.error(f"Nie udało się zapisać belki: {e}")
