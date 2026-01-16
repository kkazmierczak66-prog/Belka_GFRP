import streamlit as st
import pandas as pd
import math
import plotly.graph_objects as go

# ==========================
# Google Sheets – receptury (BEZ ZMIAN)
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


@st.cache_data(show_spinner=False)
def read_recipes_summary(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """
    Czyta arkusz receptur i zwraca wiersze PODSUMOWAŃ (SUMMARY / __SUMMARY__),
    zwracając DataFrame z kolumnami:
    recipe_name, gestosc_mix_kgm3, fck_mpa, fctm_mpa, ecm_gpa
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


# ---------------------------
# USTAWIENIA STRONY
# ---------------------------
st.set_page_config(page_title="Belka GFRP", layout="wide")
title_col, refresh_col = st.columns([10, 2])

with title_col:
    st.title("Definicja belki i obliczenia")

with refresh_col:
    if st.button(
        "↻ Odśwież dane",
        use_container_width=True,
        help="Wymusza ponowne pobranie danych z Google Sheets i przeliczenie widoku",
    ):
        st.cache_data.clear()

        for k in [
            "df_beams_i",
            "df_beams_tpd",
            "df_tests_all",
            "df_exec_all",
            "df_gfrp",
            "df_recipes",
            "df_materials",
        ]:
            if k in st.session_state:
                del st.session_state[k]

        for k in list(st.session_state.keys()):
            if isinstance(k, str) and (k.startswith("exec__") or k.startswith("tests__")):
                del st.session_state[k]

        st.rerun()

st.markdown("---")
# ============================================================
# LOAD BELKI I Z BAZY (Google Sheets)
# ============================================================

SHEET_BEAMS_I = st.secrets.get("SHEET_BEAMS_I", "belki i")

def _to_float(x, default=None):
    try:
        if x is None:
            return default
        s = str(x).strip()
        if s == "":
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
        if s == "":
            return default
        return int(float(s.replace(",", ".")))
    except Exception:
        return default

@st.cache_data(show_spinner=False)
def read_beams_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Czyta arkusz 'belki i' do DataFrame (z formułami)."""
    if not GS_RECIPES_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account).")

    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    try:
        ws = ss.worksheet(sheet_name)
    except Exception:
        raise RuntimeError(f"Nie znaleziono zakładki „{sheet_name}” w pliku Google Sheets.")

    if get_as_dataframe is not None:
        df = get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
    else:
        rows = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(rows)

    return df


# --- mechanizm "payload defaults" żeby nie było warningów i APIException ---
LOAD_KEYS = [
    # beton
    "beton_mode", "beton_recipe_name", "beton_fck", "beton_fctm", "beton_rho", "beton_Ec",
    # geometria + limity
    "geom_L", "geom_b", "geom_h", "lim_mmin", "lim_mmax",
    # gfrp selectbox
    "rebar_gfrp_sel",
    # zbrojenie dolne
    "z_ot_dolna", "z_ot_gorna", "z_ot_boczna", "z_odst_poziomy", "z_odst_pionowy", "z_n_wlasne", "z_warstwy_wlasne",
    # zbrojenie górne
    "s_ot_gorna", "s_ot_dolna", "s_ot_boczna", "s_odst_poziomy", "s_odst_pionowy", "s_n_wlasne", "s_warstwy_wlasne",
    # ścinanie
    "shear_choice_mode", "shear_P_custom_kN",
    # zapis
    "beam_name_to_save", "chk_overwrite_beam",
]

# ✅ helper do domyślnych wartości z payloadu
def d(key: str, default):
    return st.session_state.get(f"__d__{key}", default)


def _apply_load_payload_if_needed():
    """
    Jeśli istnieje payload ładowania, to:
    1) kasujemy klucze widgetów (żeby Streamlit nie trzymał starych wartości),
    2) ustawiamy BOTH:
       - __d__{key} (dla Twojej logiki defaultów)
       - {key} w session_state (żeby WIDGETY faktycznie przyjęły wartości)
    """
    payload = st.session_state.pop("__load_payload_i__", None)
    if not payload:
        return

    # 1) usuń stan widgetów (żeby mogły przyjąć nowe wartości)
    for k in LOAD_KEYS:
        if k in st.session_state:
            del st.session_state[k]

    # 2) ustaw domyślne + faktyczne wartości widgetów
    for k, v in payload.items():
        # zostawiamy "specjalne" klucze tylko w __d__ (np. __gfrp_bar_id__)
        st.session_state[f"__d__{k}"] = v

        # TYLKO dla normalnych kluczy widgetów ustawiamy stan widgetu
        if isinstance(k, str) and not k.startswith("__"):
            st.session_state[k] = v

    # 3) sekcja zapisu — ma się od razu pokazać nazwa w input
    st.session_state["beam_name_to_save"] = payload.get("beam_name_to_save", "")
    st.session_state["chk_overwrite_beam"] = bool(payload.get("chk_overwrite_beam", False))



def _request_widget_reset(widget_key: str, placeholder_value: str):
    st.session_state[f"__reset__{widget_key}"] = placeholder_value

def _apply_widget_reset_if_needed(widget_key: str):
    flag_key = f"__reset__{widget_key}"
    if flag_key in st.session_state:
        st.session_state[widget_key] = st.session_state[flag_key]
        del st.session_state[flag_key]


def _build_payload_from_row(row: dict) -> dict:
    """
    Mapowanie nagłówków INPUT_* -> klucze widgetów w Twoim kodzie.
    """
    mode_raw = str(row.get("INPUT_beton_mode", "")).strip().lower()
    if mode_raw in ("gsheet", "google", "baza", "db"):
        beton_mode_label = "Wybór z bazy danych"
    elif mode_raw in ("manual", "recznie", "ręcznie"):
        beton_mode_label = "Definiuj ręcznie (brak obliczeń punktacji)"
    else:
        beton_mode_label = "Wybór z bazy danych"

    payload = {
        # beton
        "beton_mode": beton_mode_label,
        "beton_recipe_name": str(row.get("INPUT_beton_recipe_name", "") or "").strip(),
        "beton_fck": _to_float(row.get("INPUT_fck_MPa"), 25.0),
        "beton_fctm": _to_float(row.get("INPUT_fctm_MPa"), 2.6),
        "beton_rho": _to_float(row.get("INPUT_rho_kgm3"), 2400.0),
        "beton_Ec": _to_float(row.get("INPUT_Ec_GPa"), 30.0),

        # geometria + limity
        "geom_L": _to_float(row.get("INPUT_L_m"), 1.0),
        "geom_b": _to_float(row.get("INPUT_b_cm"), 10.0),
        "geom_h": _to_float(row.get("INPUT_h_cm"), 20.0),
        "lim_mmin": _to_float(row.get("INPUT_masa_min_kg"), 5.0),
        "lim_mmax": _to_float(row.get("INPUT_masa_max_kg"), 15.0),

        # zbrojenie dolne (z_*)
        "z_ot_dolna": _to_float(row.get("INPUT_z_ot_dolna_mm"), 5.0),
        "z_ot_gorna": _to_float(row.get("INPUT_z_ot_gorna_mm"), 5.0),
        "z_ot_boczna": _to_float(row.get("INPUT_z_ot_boczna_mm"), 5.0),
        "z_odst_poziomy": _to_float(row.get("INPUT_z_odst_poziomy_mm"), 5.0),
        "z_odst_pionowy": _to_float(row.get("INPUT_z_odst_pionowy_mm"), 5.0),
        "z_n_wlasne": _to_int(row.get("INPUT_z_n_wlasne"), 4),
        "z_warstwy_wlasne": _to_int(row.get("INPUT_z_warstwy_wlasne"), 1),

        # zbrojenie górne (s_*)
        "s_ot_gorna": _to_float(row.get("INPUT_s_ot_gorna_mm"), 5.0),
        "s_ot_dolna": _to_float(row.get("INPUT_s_ot_dolna_mm"), 5.0),
        "s_ot_boczna": _to_float(row.get("INPUT_s_ot_boczna_mm"), 5.0),
        "s_odst_poziomy": _to_float(row.get("INPUT_s_odst_poziomy_mm"), 5.0),
        "s_odst_pionowy": _to_float(row.get("INPUT_s_odst_pionowy_mm"), 5.0),
        "s_n_wlasne": _to_int(row.get("INPUT_s_n_wlasne"), 2),
        "s_warstwy_wlasne": _to_int(row.get("INPUT_s_warstwy_wlasne"), 1),

        # ścinanie
        "shear_choice_mode": str(row.get("INPUT_shear_choice_mode", "") or "").strip(),
        "shear_P_custom_kN": _to_float(row.get("INPUT_shear_P_custom_kN"), 0.0),

        # zapis
        "beam_name_to_save": str(row.get("Nazwa belki", "") or "").strip(),
        "chk_overwrite_beam": True,
    }

    # pręt GFRP: w arkuszu jest ID, a w selectbox masz index df_gfrp -> rozwiązanie indexem niżej
    payload["__gfrp_bar_id__"] = _to_int(row.get("INPUT_gfrp_bar_id"), None)

    return payload


def load_beam_i_ui():
    st.subheader("Wczytaj belkę z bazy")

    try:
        df = read_beams_sheet(SPREADSHEET_ID, SHEET_BEAMS_I)
    except Exception as e:
        st.error(f"Nie udało się wczytać arkusza „{SHEET_BEAMS_I}”: {e}")
        return

    if df is None or df.empty:
        st.info(f"Arkusz „{SHEET_BEAMS_I}” jest pusty.")
        return

    if "ID" in df.columns:
        df["_ID_num"] = pd.to_numeric(df["ID"], errors="coerce")
        df = df.sort_values("_ID_num", ascending=True)

    records = df.to_dict(orient="records")

    placeholder = "— wybierz belkę —"
    options = [placeholder] + [
        f"ID={r.get('ID','?')} • {r.get('Nazwa belki','(bez nazwy)')}"
        for r in records
    ]

    key_sel = "beam_i_to_load"
    _apply_widget_reset_if_needed(key_sel)

    if key_sel not in st.session_state:
        st.session_state[key_sel] = placeholder

    colA, colB = st.columns([8, 2], vertical_alignment="top")

    with colA:
        choice = st.selectbox("", options=options, key=key_sel)

    with colB:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        clicked = st.button("📥 Wczytaj", use_container_width=True, disabled=(choice == placeholder))

    if clicked and choice != placeholder:
        idx = options.index(choice) - 1
        row = records[idx]

        payload = _build_payload_from_row(row)

        st.session_state["__load_payload_i__"] = payload
        _request_widget_reset(key_sel, placeholder)

        st.success(f"Wczytano belkę: {row.get('Nazwa belki','')}")
        st.rerun()


_apply_load_payload_if_needed()
load_beam_i_ui()

st.markdown("---")
# ============================================================
# SEKCJA – DANE DOTYCZĄCE BETONU (BEZ ZMIAN)
# ============================================================
st.header("Dane dotyczące betonu")

col_b1, col_b2 = st.columns([1, 2])

with col_b1:
    beton_opts = ["Wybór z bazy danych", "Definiuj ręcznie (brak obliczeń punktacji)"]
    beton_def = st.session_state.get("__d__beton_mode", beton_opts[0])
    beton_idx = beton_opts.index(beton_def) if beton_def in beton_opts else 0

    beton_mode = st.radio(
        "Sposób wprowadzania danych:",
        beton_opts,
        index=beton_idx,
        key="beton_mode",
    )

with col_b2:
    if beton_mode == "Wybór z bazy danych":
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
                recipe_names = sorted(df_recipes["recipe_name"].dropna().unique().tolist())
                sel_default = st.session_state.get("__d__beton_recipe_name", None)
                if sel_default in recipe_names:
                    idx_default = recipe_names.index(sel_default)
                else:
                    idx_default = 0

                sel_recipe = st.selectbox(
                    "Receptura z bazy danych:",
                    recipe_names,
                    index=idx_default,
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

    else:
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            f_ck_manual = st.number_input(
                "f_ck [MPa]",
                min_value=5.0,
                value=float(st.session_state.get("__d__beton_fck", 25.0)),
                step=1.0,
                key="beton_fck",
            )
        with col_m2:
            f_ctm_manual = st.number_input(
                "f_ctm [MPa]",
                min_value=1.0,
                value=float(st.session_state.get("__d__beton_fctm", 2.6)),
                step=0.1,
                key="beton_fctm",
            )
        with col_m3:
            rho_manual = st.number_input(
                "ρ mieszanki [kg/m³]",
                min_value=1000.0,
                value=float(st.session_state.get("__d__beton_rho", 2400.0)),
                step=50.0,
                key="beton_rho",
            )
        with col_m4:
            E_c_manual = st.number_input(
                "E_c [GPa]",
                min_value=10.0,
                value=float(st.session_state.get("__d__beton_Ec", 30.0)),
                step=1.0,
                key="beton_Ec",
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

# ============================================================
# SEKCJA – PARAMETRY BELKI (PROSTOKĄT)
# ============================================================
beton_info = st.session_state.get("beton_dane", {}) or {}

rho = beton_info.get("rho", 2400.0)
try:
    rho = float(rho) if rho is not None and not math.isnan(float(rho)) else 2400.0
except Exception:
    rho = 2400.0

st.header("Parametry belki")

col_g1, col_g2, col_g3 = st.columns(3)
with col_g1:
    L_beam = st.number_input(
        "Długość belki L [m]",
        min_value=0.01,
        value=float(st.session_state.get("__d__geom_L", 1.0)),
        step=0.1,
        key="geom_L",
    )
with col_g2:
    b_rect = st.number_input(
        "Szerokość przekroju b [cm]",
        min_value=0.0,
        value=float(st.session_state.get("__d__geom_b", 10.0)),
        step=0.5,
        key="geom_b",
    )
with col_g3:
    h_rect = st.number_input(
        "Wysokość przekroju h [cm]",
        min_value=0.0,
        value=float(st.session_state.get("__d__geom_h", 20.0)),
        step=0.5,
        key="geom_h",
    )

st.subheader("Ograniczenia")

col_lim1, col_lim2 = st.columns(2)
with col_lim1:
    masa_min = st.number_input(
        "Minimalna masa belki [kg]",
        min_value=0.0,
        value=float(st.session_state.get("__d__lim_mmin", 5.0)),
        step=0.5,
        key="lim_mmin",
    )
with col_lim2:
    masa_max = st.number_input(
        "Maksymalna masa belki [kg]",
        min_value=0.0,
        value=float(st.session_state.get("__d__lim_mmax", 15.0)),
        step=0.5,
        key="lim_mmax",
    )

st.markdown("---")

# ============================================================
# OBLICZENIA GEOMETRII (PROSTOKĄT)
# ============================================================
b_m = b_rect / 100.0
h_m = h_rect / 100.0

A_min = masa_min / (rho * L_beam) if rho > 0 and L_beam > 0 else 0.0
A_max = masa_max / (rho * L_beam) if rho > 0 and L_beam > 0 else 0.0

A = b_m * h_m
y_c = h_m / 2.0 if A > 0 else 0.0
I = (b_m * h_m**3) / 12.0 if A > 0 else 0.0
I_do_A = (I / A) if A > 0 else 0.0

A_cm2, A_min_cm2, A_max_cm2 = A * 1e4, A_min * 1e4, A_max * 1e4
I_cm4, I_do_A_cm2, y_c_cm = I * 1e8, I_do_A * 1e4, y_c * 100.0

masa_belki = rho * A * L_beam  # kg

if A > 0 and A_min <= A <= A_max and A_min < A_max:
    status_text, status_color = "✅ Pole przekroju mieści się w zakresie z masy.", "lime"
elif A == 0:
    status_text, status_color = "⚠️ Pole przekroju wynosi 0. Zwiększ wymiary.", "orange"
elif A_min >= A_max:
    status_text, status_color = "⚠️ Zakres mas jest niespójny (masa_min ≥ masa_max).", "orange"
else:
    status_text, status_color = "❌ Pole przekroju poza zakresem wynikającym z masy.", "red"

col_l, col_r = st.columns([1, 1])
with col_l:
    st.subheader("Obliczenia — geometria (prostokąt)")
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

with col_r:
    import plotly.graph_objects as go

    b_cm = float(b_rect)
    h_cm = float(h_rect)
    y_c_cm = float(y_c_cm)

    pad_x = max(1.0, 0.15 * b_cm)
    pad_y = max(1.0, 0.20 * h_cm)

    x_min, x_max = -b_cm / 2.0 - pad_x, b_cm / 2.0 + pad_x
    y_min, y_max = -pad_y, h_cm + pad_y

    fig = go.Figure()
    fig.update_layout(
        plot_bgcolor="black",
        paper_bgcolor="black",
        font=dict(color="white"),
        width=None,
        height=650,
        margin=dict(l=10, r=160, t=40, b=10),
        showlegend=False,
    )

    if b_cm > 0 and h_cm > 0:
        fig.add_shape(
            type="rect",
            x0=-b_cm / 2.0,
            y0=0,
            x1=b_cm / 2.0,
            y1=h_cm,
            line=dict(color="white", width=2),
            fillcolor="rgba(100,160,255,0.6)",
        )

    if b_cm > 0 and h_cm > 0:
        fig.add_shape(
            type="line",
            x0=x_min,
            y0=y_c_cm,
            x1=x_max,
            y1=y_c_cm,
            line=dict(color="red", width=2, dash="dash"),
        )

        fig.add_annotation(
            x=0.65,
            xref="paper",
            y=y_c_cm,
            yref="y",
            text=f"y_c = {y_c_cm:.2f} cm",
            showarrow=False,
            font=dict(color="red"),
            xanchor="left",
            align="left",
        )

    def add_dim_h(fig, x0, x1, y, tick=0.6, w=2):
        fig.add_shape(type="line", x0=x0, y0=y, x1=x1, y1=y, line=dict(color="white", width=w))
        fig.add_shape(type="line", x0=x0, y0=y - tick, x1=x0, y1=y + tick, line=dict(color="white", width=w))
        fig.add_shape(type="line", x0=x1, y0=y - tick, x1=x1, y1=y + tick, line=dict(color="white", width=w))

    def add_dim_v(fig, x, y0, y1, tick=0.6, w=2):
        fig.add_shape(type="line", x0=x, y0=y0, x1=x, y1=y1, line=dict(color="white", width=w))
        fig.add_shape(type="line", x0=x - tick, y0=y0, x1=x + tick, y1=y0, line=dict(color="white", width=w))
        fig.add_shape(type="line", x0=x - tick, y0=y1, x1=x + tick, y1=y1, line=dict(color="white", width=w))

    tick_h = max(0.4, 0.03 * b_cm)
    tick_v = max(0.4, 0.03 * h_cm)

    y_dim_b = y_min + 0.8
    add_dim_h(fig, -b_cm / 2.0, b_cm / 2.0, y_dim_b, tick=tick_v)

    fig.add_annotation(
        x=0,
        y=y_dim_b,
        text=f"b = {b_cm:g} cm",
        showarrow=False,
        font=dict(color="white"),
        xanchor="center",
        yanchor="top",
        yshift=-16,
    )

    x_dim_h = x_max - 0.5
    add_dim_v(fig, x_dim_h, 0, h_cm, tick=tick_h)
    fig.add_annotation(
        x=x_dim_h + 0.6,
        y=h_cm / 2.0,
        text=f"h = {h_cm:g} cm",
        showarrow=False,
        textangle=90,
        font=dict(color="white"),
    )

    fig.update_xaxes(
        range=[x_min, x_max],
        title="Szerokość [cm], oś x (0 w środku przekroju)",
        gridcolor="#444",
        zerolinecolor="#666",
    )
    fig.update_yaxes(
        range=[y_min, y_max],
        title="Wysokość [cm], oś y (0 = dół przekroju)",
        gridcolor="#444",
        zerolinecolor="#666",
        scaleanchor="x",
        scaleratio=1,
    )

    st.plotly_chart(fig, use_container_width=True)


# ==========================
# Google Sheets – baza prętów GFRP
# ==========================
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

    for c in ["id", "srednica_mm", "R_t_MPa", "E_GPa", "τ_base_MPa", "gestosc_gcm3", "cena_pln", "co2e_kgkg"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["profil"] = df["profil"].astype(object)
    df["cena_za"] = df["cena_za"].astype(object)

    return df[wanted]


# ============================================================
# SEKCJA ZBROJENIE — prostokąt
# ============================================================
import plotly.graph_objects as go

st.markdown("---")
st.header("Zbrojenie")

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

if df_gfrp is None or df_gfrp.empty:
    st.warning("Baza prętów GFRP jest pusta. Uzupełnij arkusz SHEET_GFRP, aby korzystać z tej sekcji.")
    st.stop()

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

# --- domyślny wybór pręta po ID z arkusza (jeśli załadowano belkę) ---
default_sel_index = 0
desired_bar_id = st.session_state.get("__d____gfrp_bar_id__", None)

if desired_bar_id is not None and "id" in df_gfrp.columns:
    m = df_gfrp[df_gfrp["id"] == desired_bar_id]
    if not m.empty:
        opt_val = m.index[0]
        if opt_val in options_idx:
            default_sel_index = options_idx.index(opt_val)

sel_idx = st.selectbox(
    "Pręt z bazy danych:",
    options_idx,
    format_func=format_bar,
    index=default_sel_index,
    key="rebar_gfrp_sel",
)

row = df_gfrp.loc[sel_idx]

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
    "phi_mm": phi_mm,
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

phi_bot_mm = phi_mm
phi_top_mm = phi_mm

b_rect_mm = float(b_rect) * 10.0
h_rect_mm = float(h_rect) * 10.0

# ============================================================
# 2) ZBROJENIE DOLNE (dół przekroju)
# ============================================================
st.subheader("Rozmieszczenie prętów — zbrojenie dolne")

col_left, col_right = st.columns([1, 1])

# ✅ POPRAWKA: value bierze z __d__ (czyli z wczytanej belki)
with col_left:
    col1, col2, col3 = st.columns(3)
    with col1:
        otulina_dolna = st.number_input("Otulina dolna [mm]", value=float(d("z_ot_dolna", 5.0)), key="z_ot_dolna")
    with col2:
        otulina_gorna = st.number_input("Otulina górna [mm]", value=float(d("z_ot_gorna", 5.0)), key="z_ot_gorna")
    with col3:
        otulina_boczna = st.number_input("Otulina boczna [mm]", value=float(d("z_ot_boczna", 5.0)), key="z_ot_boczna")

    col4, col5 = st.columns(2)
    with col4:
        odleglosc_pozioma = st.number_input(
            "Odstęp poziomy między prętami [mm] (clear)",
            value=float(d("z_odst_poziomy", 5.0)),
            key="z_odst_poziomy",
        )
    with col5:
        odleglosc_pionowa = st.number_input(
            "Odstęp pionowy między warstwami [mm] (clear)",
            value=float(d("z_odst_pionowy", 5.0)),
            key="z_odst_pionowy",
        )

    n_wlasne = int(st.number_input("Liczba prętów w 1 warstwie [szt.]", min_value=0, value=int(d("z_n_wlasne", 4)), step=1, key="z_n_wlasne"))
    warstwy_wlasne = int(st.number_input("Liczba warstw [szt.]", min_value=0, value=int(d("z_warstwy_wlasne", 1)), step=1, key="z_warstwy_wlasne"))

    szer_dostepna = b_rect_mm - 2 * otulina_boczna
    wysokosc_dostepna = h_rect_mm - otulina_dolna - otulina_gorna

    req_width = phi_bot_mm * n_wlasne + odleglosc_pozioma * max(0, n_wlasne - 1)
    req_height = phi_bot_mm * warstwy_wlasne + odleglosc_pionowa * max(0, warstwy_wlasne - 1)

    violations = []
    if szer_dostepna <= 0:
        violations.append("❌ Brak światła na szerokości (otuliny boczne zjadają cały przekrój).")
    if wysokosc_dostepna <= 0:
        violations.append("❌ Brak światła na wysokości (otulina górna/dolna zjada cały przekrój).")
    if n_wlasne > 0 and warstwy_wlasne > 0:
        if req_width > max(0.0, szer_dostepna):
            violations.append("❌ Pręty **nie mieszczą się na szerokość** dla zadanych otulin i odstępów.")
        if req_height > max(0.0, wysokosc_dostepna):
            violations.append("❌ Pręty/warstwy **nie mieszczą się na wysokość** dla zadanych otulin i odstępów.")

    fits_bot = (
        n_wlasne > 0
        and warstwy_wlasne > 0
        and szer_dostepna > 0
        and wysokosc_dostepna > 0
        and req_width <= szer_dostepna
        and req_height <= wysokosc_dostepna
    )

    for v in violations:
        st.error(v)

    A_pręt = math.pi * (phi_bot_mm / 2.0) ** 2 / 100.0  # cm²
    As_bot = A_pręt * n_wlasne * warstwy_wlasne
    O_bot = math.pi * phi_bot_mm * n_wlasne * warstwy_wlasne
    n_bot = n_wlasne * warstwy_wlasne

    st.session_state["bottom_rebar_summary"] = {
        "As_cm2": float(As_bot),
        "n_bars": int(n_bot),
        "perimeter_mm": float(O_bot),
    }

with col_right:
    fig_bot = go.Figure()
    fig_bot.update_layout(
        plot_bgcolor="black",
        paper_bgcolor="black",
        font=dict(color="white"),
        width=None,
        height=500,
        showlegend=False,
        margin=dict(l=10, r=10, t=10, b=10),
    )

    fig_bot.add_shape(
        type="rect",
        x0=0,
        y0=0,
        x1=b_rect_mm,
        y1=h_rect_mm,
        line=dict(color="white", width=2),
        fillcolor="rgba(80,80,80,0.35)",
    )

    if not fits_bot and n_wlasne > 0 and warstwy_wlasne > 0:
        st.warning("⚠️ Zbrojenie dolne nie mieści się geometrycznie — nie jest rysowane na przekroju.")

    if fits_bot:
        fi = phi_bot_mm

        szer_dostepna_vis = max(0.0, b_rect_mm - 2 * otulina_boczna)
        if n_wlasne == 1:
            x_poz = [b_rect_mm / 2.0]
        else:
            x0 = otulina_boczna + fi / 2.0
            x_poz = [x0 + i * (szer_dostepna_vis - fi) / (n_wlasne - 1) for i in range(n_wlasne)]

        y0 = otulina_dolna + fi / 2.0
        krok = fi + odleglosc_pionowa
        y_poz = [y0 + j * krok for j in range(warstwy_wlasne)]

        for y in y_poz:
            for x in x_poz:
                fig_bot.add_shape(
                    type="circle",
                    x0=x - fi / 2.0, y0=y - fi / 2.0,
                    x1=x + fi / 2.0, y1=y + fi / 2.0,
                    line=dict(color="red", width=1.5),
                    fillcolor="red",
                )

    fig_bot.update_xaxes(
        title="Szerokość przekroju [mm]",
        range=[-10, b_rect_mm + 10],
        scaleanchor="y",
        scaleratio=1,
    )
    fig_bot.update_yaxes(
        title="Wysokość przekroju [mm] (0 = dół)",
        range=[-10, h_rect_mm + 10],
    )
    st.plotly_chart(fig_bot, use_container_width=True)

# ============================================================
# 3) ZBROJENIE GÓRNE (góra przekroju)
# ============================================================
st.subheader("Rozmieszczenie prętów — zbrojenie górne")

col_ws_left, col_ws_right = st.columns([1, 1])

# ✅ POPRAWKA: value bierze z __d__ (czyli z wczytanej belki)
with col_ws_left:
    colw1, colw2, colw3 = st.columns(3)
    with colw1:
        otulina_gorna_s = st.number_input("Otulina górna [mm] (od góry przekroju)", value=float(d("s_ot_gorna", 5.0)), key="s_ot_gorna")
    with colw2:
        otulina_dolna_s = st.number_input("Otulina dolna [mm] (od dołu przekroju)", value=float(d("s_ot_dolna", 5.0)), key="s_ot_dolna")
    with colw3:
        otulina_boczna_s = st.number_input("Otulina boczna [mm]", value=float(d("s_ot_boczna", 5.0)), key="s_ot_boczna")

    colw4, colw5 = st.columns(2)
    with colw4:
        odleglosc_pozioma_s = st.number_input("Odstęp poziomy między prętami [mm] (clear)", value=float(d("s_odst_poziomy", 5.0)), key="s_odst_poziomy")
    with colw5:
        odleglosc_pionowa_s = st.number_input("Odstęp pionowy między warstwami [mm] (clear)", value=float(d("s_odst_pionowy", 5.0)), key="s_odst_pionowy")

    n_wlasne_s = int(st.number_input("Liczba prętów w 1 warstwie [szt.]", min_value=0, value=int(d("s_n_wlasne", 2)), step=1, key="s_n_wlasne"))
    warstwy_wlasne_s = int(st.number_input("Liczba warstw [szt.]", min_value=0, value=int(d("s_warstwy_wlasne", 1)), step=1, key="s_warstwy_wlasne"))

    szer_dostepna_s = b_rect_mm - 2 * otulina_boczna_s
    wysokosc_dostepna_s = h_rect_mm - otulina_gorna_s - otulina_dolna_s

    req_width_s = phi_top_mm * n_wlasne_s + odleglosc_pozioma_s * max(0, n_wlasne_s - 1)
    req_height_s = phi_top_mm * warstwy_wlasne_s + odleglosc_pionowa_s * max(0, warstwy_wlasne_s - 1)

    violations_s = []
    if szer_dostepna_s <= 0:
        violations_s.append("❌ Brak światła na szerokości (otuliny boczne zjadają cały przekrój).")
    if wysokosc_dostepna_s <= 0:
        violations_s.append("❌ Brak światła na wysokości (otulina górna/dolna zjada cały przekrój).")
    if n_wlasne_s > 0 and warstwy_wlasne_s > 0:
        if req_width_s > max(0.0, szer_dostepna_s):
            violations_s.append("❌ Pręty **nie mieszczą się na szerokość** dla zadanych otulin i odstępów.")
        if req_height_s > max(0.0, wysokosc_dostepna_s):
            violations_s.append("❌ Pręty/warstwy **nie mieszczą się na wysokość** dla zadanych otulin i odstępów.")

    fits_top = (
        n_wlasne_s > 0
        and warstwy_wlasne_s > 0
        and szer_dostepna_s > 0
        and wysokosc_dostepna_s > 0
        and req_width_s <= max(0.0, szer_dostepna_s)
        and req_height_s <= max(0.0, wysokosc_dostepna_s)
    )

    for v in violations_s:
        st.error(v)

    A_pręt_s = math.pi * (phi_top_mm / 2.0) ** 2 / 100.0  # cm²
    As_top = A_pręt_s * n_wlasne_s * warstwy_wlasne_s
    O_top = math.pi * phi_top_mm * n_wlasne_s * warstwy_wlasne_s
    n_top = n_wlasne_s * warstwy_wlasne_s

    st.session_state["top_rebar_summary"] = {
        "As_cm2": float(As_top),
        "n_bars": int(n_top),
        "perimeter_mm": float(O_top),
    }

with col_ws_right:
    fig_top = go.Figure()
    fig_top.update_layout(
        plot_bgcolor="black",
        paper_bgcolor="black",
        font=dict(color="white"),
        width=None,
        height=500,
        showlegend=False,
        margin=dict(l=10, r=10, t=10, b=10),
    )

    fig_top.add_shape(
        type="rect",
        x0=0,
        y0=0,
        x1=b_rect_mm,
        y1=h_rect_mm,
        line=dict(color="white", width=2),
        fillcolor="rgba(80,80,80,0.35)",
    )

    if not fits_top and n_wlasne_s > 0 and warstwy_wlasne_s > 0:
        st.warning("⚠️ Zbrojenie górne nie mieści się geometrycznie — nie jest rysowane na przekroju.")

    if fits_top:
        fi = phi_top_mm
        szer_dostepna_vis = max(0.0, b_rect_mm - 2 * otulina_boczna_s)

        if n_wlasne_s == 1:
            x_poz_s = [b_rect_mm / 2.0]
        else:
            x0_s = otulina_boczna_s + fi / 2.0
            x_poz_s = [x0_s + i * (szer_dostepna_vis - fi) / (n_wlasne_s - 1) for i in range(n_wlasne_s)]

        y_top_center = h_rect_mm - otulina_gorna_s - fi / 2.0
        krok = fi + odleglosc_pionowa_s
        y_poz_s = [y_top_center - j * krok for j in range(warstwy_wlasne_s)]

        min_center = otulina_dolna_s + fi / 2.0
        y_poz_s = [y for y in y_poz_s if y >= min_center]

        for y in y_poz_s:
            for x in x_poz_s:
                fig_top.add_shape(
                    type="circle",
                    x0=x - fi / 2.0, y0=y - fi / 2.0,
                    x1=x + fi / 2.0, y1=y + fi / 2.0,
                    line=dict(color="red", width=1.5),
                    fillcolor="red",
                )

    fig_top.update_xaxes(
        title="Szerokość przekroju [mm]",
        range=[-10, b_rect_mm + 10],
        scaleanchor="y",
        scaleratio=1,
    )
    fig_top.update_yaxes(
        title="Wysokość przekroju [mm] (0 = dół)",
        range=[-10, h_rect_mm + 10],
    )
    st.plotly_chart(fig_top, use_container_width=True)

# ============================================================
# 4) PODSUMOWANIE ZBROJENIA
# ============================================================
st.subheader("Podsumowanie zbrojenia")

try:
    As_bot_v = float(As_bot)
    n_bot_v = int(n_bot)
    O_bot_v = float(O_bot)
except Exception:
    As_bot_v, n_bot_v, O_bot_v = 0.0, 0, 0.0

try:
    As_top_v = float(As_top)
    n_top_v = int(n_top)
    O_top_v = float(O_top)
except Exception:
    As_top_v, n_top_v, O_top_v = 0.0, 0, 0.0

As_total = As_bot_v + As_top_v
n_total = n_bot_v + n_top_v
O_total = O_bot_v + O_top_v

df_rebar_summary = pd.DataFrame(
    [
        ["Zbrojenie dolne", f"{As_bot_v:.2f}", n_bot_v, f"{O_bot_v:.0f}"],
        ["Zbrojenie górne", f"{As_top_v:.2f}", n_top_v, f"{O_top_v:.0f}"],
        ["Łącznie",         f"{As_total:.2f}", n_total, f"{O_total:.0f}"],
    ],
    columns=["Element", "As [cm²]", "Liczba prętów", "Obwód [mm]"],
)

st.table(df_rebar_summary.set_index("Element"))



# ============================================================
# SEKCJA – WYTRZYMAŁOŚĆ NA ŚCINANIE (PROSTOKĄT)
# ============================================================

st.markdown("---")
st.header("Wytrzymałość na ścinanie")

# ---- pomocnicze funkcje dla procedur (bez zmian) ----
def shear_capacity_ACI(bw_mm, d_mm, Af_mm2, Ef_GPa, Ec_GPa, fck_MPa) -> float:
    """Nośność na ścinanie wg ACI 440 (bez współczynników bezpieczeństwa). Zwraca Pmax [kN]."""
    try:
        if bw_mm <= 0 or d_mm <= 0 or fck_MPa <= 0 or Af_mm2 <= 0 or Ef_GPa <= 0 or Ec_GPa <= 0:
            return float("nan")

        rho_f = Af_mm2 / (bw_mm * d_mm)
        eta_f = Ef_GPa / Ec_GPa

        k = math.sqrt(max(0.0, 2.0 * eta_f * rho_f + (eta_f * rho_f) ** 2))
        Vc_N = (2.0 / 5.0) * math.sqrt(fck_MPa) * bw_mm * d_mm * k
        Pmax_kN = 2.0 * Vc_N / 1000.0  # dwa obciążenia P/2
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

        lambda_c = 1.0
        phi_c = 1.0

        dv_mm = min(0.9 * d_mm, 0.72 * h_mm)
        rho_l = Af_mm2 / (bw_mm * d_mm)

        kr = 1.0 + (Ef_GPa * rho_l) ** (1.0 / 3.0)

        km = math.sqrt(d_mm / dv_mm) if dv_mm > 0 else 0.0
        km = min(km, 1.0)

        Vc_raw_N = (
            0.05 * lambda_c * phi_c * km * kr
            * (fck_MPa ** (1.0 / 3.0))
            * bw_mm * dv_mm
        )

        Vc_min_N = 0.11 * phi_c * math.sqrt(fck_MPa) * bw_mm * dv_mm
        Vc_max_N = 0.22 * phi_c * math.sqrt(fck_MPa) * bw_mm * dv_mm

        Vc_N = min(max(Vc_raw_N, Vc_min_N), Vc_max_N)

        Pmax_kN = 2.0 * Vc_N / 1000.0
        return Pmax_kN
    except Exception:
        return float("nan")


# ---- właściwe obliczenia sekcji ----
try:
    # Geometria (prostokąt) w mm
    bw_mm = float(b_rect) * 10.0
    H_mm  = float(h_rect) * 10.0

    if bw_mm <= 0 or H_mm <= 0:
        st.warning("⚠️ Podaj poprawne wymiary przekroju (b, h) > 0.")
        st.stop()

    # Zbrojenie: bierzemy pręt z bazy (średnica, E)
    rebar_info = st.session_state.get("rebar_bar", {}) or {}
    if not rebar_info:
        st.warning("⚠️ Wybierz pręt z bazy GFRP w sekcji Zbrojenie.")
        st.stop()

    phi_mm = float(rebar_info.get("phi_mm", float("nan")))
    Ef_GPa = float(rebar_info.get("E_GPa", float("nan")))

    if not (phi_mm == phi_mm) or phi_mm <= 0:
        st.warning("⚠️ Brak poprawnej średnicy pręta.")
        st.stop()

    if not (Ef_GPa == Ef_GPa) or Ef_GPa <= 0:
        Ef_GPa = 50.0  # sensowny default

    # Beton
    beton_info = st.session_state.get("beton_dane", {}) or {}
    fck_MPa = float(beton_info.get("f_ck", float("nan")))
    Ec_GPa  = float(beton_info.get("E_c_GPa", float("nan")))

    if not (fck_MPa == fck_MPa) or fck_MPa <= 0:
        fck_MPa = 30.0
    if not (Ec_GPa == Ec_GPa) or Ec_GPa <= 0:
        Ec_GPa = 30.0

    # Liczba prętów w strefie rozciąganej (DOLNE zbrojenie: z_*)
    n_w = int(st.session_state.get("z_n_wlasne", 0))
    n_layers = int(st.session_state.get("z_warstwy_wlasne", 0))
    n_bot = max(0, n_w) * max(0, n_layers)

    if n_bot <= 0:
        st.warning("⚠️ Zdefiniuj dolne zbrojenie (liczba prętów i warstw) – inaczej nie policzę d i Af.")
        st.stop()

    # Otulina dolna + odstęp pionowy (mm) dla dolnego zbrojenia
    otulina_dolna_mm = float(st.session_state.get("z_ot_dolna", 5.0))
    odst_pion_mm     = float(st.session_state.get("z_odst_pionowy", 5.0))

    # Środek zbrojenia dolnego (średnia po warstwach) mierzony od DOŁU
    # warstwa 1: center = otulina_dolna + phi/2
    y1 = otulina_dolna_mm + phi_mm / 2.0
    y2 = y1 + (n_layers - 1) * (phi_mm + odst_pion_mm)
    y_cent = 0.5 * (y1 + y2)

    # Efektywna wysokość d = od góry przekroju do środka zbrojenia dolnego
    d_mm = H_mm - y_cent

    if d_mm <= 0:
        st.warning("⚠️ Wyszło d ≤ 0 (zbrojenie jest „powyżej” góry przekroju). Sprawdź otuliny/warstwy/φ oraz h.")
        st.stop()

    # Pole zbrojenia rozciąganego (mm²)
    Af_mm2 = n_bot * (math.pi * (phi_mm ** 2) / 4.0)

    # --- obliczenia ---
    P_ACI_kN  = shear_capacity_ACI(bw_mm, d_mm, Af_mm2, Ef_GPa, Ec_GPa, fck_MPa)
    P_JSCE_kN = shear_capacity_JSCE(bw_mm, d_mm, Af_mm2, Ef_GPa, fck_MPa)
    P_CSA_kN  = shear_capacity_CSA(bw_mm, H_mm, d_mm, Af_mm2, Ef_GPa, fck_MPa)

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
    P_list = [x for x in [P_ACI_kN, P_JSCE_kN, P_CSA_kN] if isinstance(x, (int, float)) and not math.isnan(x) and x > 0]
    P_default = min(P_list) if P_list else 0.0

    col_choice, col_custom = st.columns([2, 1])

    with col_choice:
        shear_opts = ["min(ACI, JSCE, CSA)", "ACI 440", "JSCE", "CSA", "Własna wartość"]
        shear_def = st.session_state.get("__d__shear_choice_mode", shear_opts[0])
        shear_idx = shear_opts.index(shear_def) if shear_def in shear_opts else 0

        shear_choice = st.radio(
            "Wybór wartości do obliczeń punktacji",
            shear_opts,
            index=shear_idx,
            horizontal=True,
            key="shear_choice_mode",
        )

    with col_custom:
        P_custom_kN = st.number_input(
            "P,max własna [kN]",
            min_value=0.0,
            value=float(st.session_state.get("__d__shear_P_custom_kN", float(P_default))),
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

    # zapis do session_state (jak wcześniej)
    st.session_state["shear_P_ACI_kN"] = float(P_ACI_kN) if P_ACI_kN == P_ACI_kN else float("nan")
    st.session_state["shear_P_JSCE_kN"] = float(P_JSCE_kN) if P_JSCE_kN == P_JSCE_kN else float("nan")
    st.session_state["shear_P_CSA_kN"] = float(P_CSA_kN) if P_CSA_kN == P_CSA_kN else float("nan")
    st.session_state["shear_P_used_kN"] = float(P_used_kN) if P_used_kN == P_used_kN else float("nan")



except Exception as e:
    st.error(f"Nie udało się policzyć nośności na ścinanie: {e}")


# ==============================================================
# FUNKCJE: koszt 1 m³ mieszanki + korekta materiałowa (%)
#   (potrzebne do sekcji "Punktacja")
# ==============================================================

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
      - jeśli cena_za zawiera 'kg' -> ilość_jedn = masa_kgm3,
      - jeśli cena_za zawiera 'l'  -> ilość_jedn = (masa_kgm3 / rho_kgm3) * 1000,
      - inaczej traktujemy jak 'kg'.

    Zwraca:
      (cena_m3, df_skladniki), gdzie df_skladniki ma kolumny:
      material, cena_za, cena_jedn, ilosc_jedn, koszt
    """
    if not GS_RECIPES_READY:
        raise RuntimeError("Brak konfiguracji Google Sheets.")

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

    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)

    # ---------- RECEPTURY ----------
    ws_rec = ss.worksheet(sheet_recipes)
    values = ws_rec.get_all_values()
    if not values:
        return 0.0, pd.DataFrame(columns=["material", "cena_za", "cena_jedn", "ilosc_jedn", "koszt"])

    header = values[0]
    df_rec = pd.DataFrame(values[1:], columns=header)

    for col in ["recipe_name", "nazwa", "material_id", "masa_kgm3"]:
        if col not in df_rec.columns:
            df_rec[col] = ""

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

    # ---------- JOIN ----------
    df_cost = df_rec.merge(df_mat_id, on="material_id", how="left")

    df_cost["masa_kgm3"] = df_cost["masa_kgm3"].fillna(0.0)
    df_cost["rho_gcm3"] = df_cost["rho_gcm3"].apply(_to_num)
    df_cost["rho_kgm3"] = df_cost["rho_gcm3"] * 1000.0

    # ---------- ILOŚĆ JEDNOSTEK ----------
    def _calc_amount(row):
        masa = float(row["masa_kgm3"])  # kg/m³
        rho  = float(row["rho_kgm3"]) if row["rho_kgm3"] == row["rho_kgm3"] else float("nan")
        unit_raw = str(row["cena_za_mat"] or "").strip().lower()
        unit = unit_raw.replace("/", "").replace(" ", "")

        # domyślnie jak kg
        if "l" in unit and "kg" not in unit:
            if rho == rho and rho > 0:
                vol_m3 = masa / rho      # m³/m³
                return vol_m3 * 1000.0   # l/m³
            return masa
        return masa

    df_cost["ilosc_jedn"] = df_cost.apply(_calc_amount, axis=1)

    # ---------- KOSZT ----------
    df_cost["koszt_m3"] = df_cost["ilosc_jedn"] * df_cost["cena_pln"]
    cena_m3 = float(df_cost["koszt_m3"].sum()) if not df_cost.empty else 0.0

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

    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)

    # ---------- RECEPTURY ----------
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

    # ---------- MATERIAŁY ----------
    ws_mat = ss.worksheet(sheet_materials)
    rows_mat = ws_mat.get_all_records(numericise_ignore=["all"])
    df_mat = pd.DataFrame(rows_mat)

    for c in ["id", "atrybut"]:
        if c not in df_mat.columns:
            df_mat[c] = None

    df_mat["id"] = df_mat["id"].apply(_strip_apos)
    df_mat["id"] = pd.to_numeric(df_mat["id"], errors="coerce")

    df_mat_small = df_mat.rename(columns={"id": "material_id"})[["material_id", "atrybut"]]

    # ---------- JOIN ----------
    df = df_rec.merge(df_mat_small, on="material_id", how="left")

    total_mass = float(df["masa_kgm3"].sum())
    s_attr = df["atrybut"].fillna("").astype(str).str.strip()

    def _mass_for(attr: str) -> float:
        return float(df.loc[s_attr == attr, "masa_kgm3"].sum())

    cement_mass = _mass_for("Cement korekta")
    flyash_mass = _mass_for("Popiół lotny dodatek")
    slag_mass   = _mass_for("Żużel wiel. dodatek")
    silica_mass = _mass_for("Pył krzem. dodatek")

    cementitious_sum = cement_mass + flyash_mass + slag_mass + silica_mass

    bonus = 0.0

    # Cement vs TOTAL
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
# SEKCJA – PUNKTACJA (PROSTOKĄT) — tak samo jak wcześniej
#   - geometria: prostokąt b,h,L
#   - zbrojenie: dół (z_*) + góra (s_*)
#   - beton + koszt mieszanki + korekty: bez zmian
#   - wynik: Cena belki netto / P,max użyte
# ============================================================

st.markdown("---")
st.header("Punktacja")

try:
    # ------------------------------------------------------
    # 1) OBJĘTOŚĆ BELKI (bez zbrojenia) — prostokąt
    # ------------------------------------------------------
    b_m = float(b_rect) / 100.0
    h_m = float(h_rect) / 100.0
    L_m = float(L_beam)

    A_m2 = b_m * h_m
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
    # 3) Korekta materiałowa [%] wg atrybutów (jak wcześniej)
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
    # 4) Korekta geometryczna [%] — zostawiamy jak było
    # ------------------------------------------------------
    geom_adj_pct = 0.0

    # ------------------------------------------------------
    # 5) PARAMETRY BETONU
    # ------------------------------------------------------
    beton_info = st.session_state.get("beton_dane", {}) or {}
    rho_conc_kgm3 = float(beton_info.get("rho", 2400.0))
    if not (rho_conc_kgm3 == rho_conc_kgm3) or rho_conc_kgm3 <= 0:
        rho_conc_kgm3 = 2400.0

    # ------------------------------------------------------
    # 6) ZBROJENIE – ilość, objętość, koszt (dół + góra)
    # ------------------------------------------------------
    rebar_info = st.session_state.get("rebar_bar", {}) or {}
    if not rebar_info:
        st.error("Nie wybrano pręta GFRP — nie mogę policzyć punktacji.")
        st.stop()

    phi_mm = float(rebar_info.get("phi_mm", float("nan")))
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

    if not (phi_mm == phi_mm) or phi_mm <= 0:
        st.error("Brak poprawnej średnicy pręta — nie mogę policzyć punktacji.")
        st.stop()

    # liczba prętów: dół (z_*) + góra (s_*)
    n_bottom = int(st.session_state.get("z_n_wlasne", 0)) * int(st.session_state.get("z_warstwy_wlasne", 0))
    n_top    = int(st.session_state.get("s_n_wlasne", 0)) * int(st.session_state.get("s_warstwy_wlasne", 0))
    n_bars_tot = int(max(0, n_bottom) + max(0, n_top))

    # geometria prętów
    area_bar_m2 = math.pi * ((phi_mm / 1000.0) ** 2) / 4.0  # mm->m
    total_length_m = n_bars_tot * L_m
    V_rebar_m3 = area_bar_m2 * total_length_m

    mass_rebar_kg = V_rebar_m3 * (rho_bar_kgm3 if (rho_bar_kgm3 == rho_bar_kgm3 and rho_bar_kgm3 > 0) else 0.0)

    # koszt zbrojenia
    if "kg" in unit_norm:
        cost_rebar_usd = price_bar * mass_rebar_kg
    elif "mb" in unit_norm or unit_norm == "m":
        cost_rebar_usd = price_bar * n_bars_tot * L_m
    else:
        # domyślnie jak "za metr"
        cost_rebar_usd = price_bar * n_bars_tot * L_m

    # ------------------------------------------------------
    # 7) CENA MIESZANKI DLA BELKI (z odjęciem objętości prętów)
    # ------------------------------------------------------
    V_conc_net_m3 = max(0.0, V_beam_m3 - V_rebar_m3)
    cost_mix_usd = mix_price_usd_m3 * V_conc_net_m3

    # ------------------------------------------------------
    # 8) MASA BELKI
    # ------------------------------------------------------
    mass_conc_kg = V_conc_net_m3 * rho_conc_kgm3
    mass_total_kg = mass_conc_kg + mass_rebar_kg

    # ------------------------------------------------------
    # 9) KOSZT TRANSPORTU – 0.01 USD za każdy rozpoczęty lb
    # ------------------------------------------------------
    lb_per_kg = 2.20462262185
    mass_lb = mass_total_kg * lb_per_kg
    transport_usd = 0.01 * math.ceil(mass_lb)

    # ------------------------------------------------------
    # 10) WYTRZYMAŁOŚĆ (P,max) — bierzemy z sekcji ścinania (jak była)
    # ------------------------------------------------------
    P_used_ext = st.session_state.get("shear_P_used_kN", None)
    P_used_kN = float(P_used_ext) if isinstance(P_used_ext, (int, float)) and P_used_ext > 0 else float("nan")

    if not (P_used_kN == P_used_kN) or P_used_kN <= 0:
        st.error("Brak poprawnej wartości P,max z sekcji 'Wytrzymałość na ścinanie'.")
        st.stop()

    # ------------------------------------------------------
    # 11) KOSZTY MATERIAŁÓW:
    #  - BRUTTO: bez korekty materiałowej
    #  - NETTO: po korekcie materiałowej (material_adj_pct jest ujemny)
    # ------------------------------------------------------
    cost_materials_brutto_usd = cost_mix_usd + cost_rebar_usd
    cost_materials_netto_usd = cost_materials_brutto_usd * (1.0 + material_adj_pct / 100.0)

    # ------------------------------------------------------
    # 12) Cena belki (brutto / netto) — jak wcześniej
    # ------------------------------------------------------
    price_beam_brutto_usd = cost_materials_brutto_usd + transport_usd

    geom_correction_usd = cost_materials_netto_usd * (geom_adj_pct / 100.0)
    price_beam_netto_usd = cost_materials_netto_usd + geom_correction_usd + transport_usd

    # ------------------------------------------------------
    # 13) WYNIK USD/kN (z ceny belki NETTO)
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

    # do zapisu belki (jak wcześniej)
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
# ZAPIS BELKI DO GOOGLE SHEETS (arkusz: "belki i")
#   - nagłówki zawsze wymuszone (_ensure_headers)
#   - w kolumnie "Geometria" zapisujemy "i"
#   - zapisujemy też INPUT_* żeby dało się odtworzyć belkę w edytorze
# ============================================================

from typing import Any, List

SHEET_BEAMS_I = st.secrets.get("SHEET_BEAMS_I", "belki i")

# ---------------------------
# Nagłówki (wyniki + INPUTS)
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

INPUT_HEADERS = [
    "INPUT_L_m",
    "INPUT_b_cm",
    "INPUT_h_cm",
    "INPUT_masa_min_kg",
    "INPUT_masa_max_kg",

    "INPUT_beton_mode",          # "gsheet" / "manual"
    "INPUT_beton_recipe_name",
    "INPUT_fck_MPa",
    "INPUT_fctm_MPa",
    "INPUT_rho_kgm3",
    "INPUT_Ec_GPa",

    "INPUT_gfrp_bar_id",
    "INPUT_phi_mm",
    "INPUT_Ef_GPa",

    # zbrojenie dolne (z_*)
    "INPUT_z_ot_dolna_mm",
    "INPUT_z_ot_gorna_mm",
    "INPUT_z_ot_boczna_mm",
    "INPUT_z_odst_poziomy_mm",
    "INPUT_z_odst_pionowy_mm",
    "INPUT_z_n_wlasne",
    "INPUT_z_warstwy_wlasne",

    # zbrojenie górne (s_*)
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

    # usuń nagłówek jeśli jest
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
# Budowanie wiersza do zapisu
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
        "",  # ID (nadamy przy zapisie)
        beam_name.strip(),
        str(recipe_name),
        "prost.",  # ✅ geometria
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

    # ---------------------------
    # INPUTS do odtworzenia belki
    # ---------------------------
    beton_mode_src = str(beton_info.get("source", ""))  # "gsheet" / "manual"
    beton_recipe_name = st.session_state.get("beton_recipe_name") or beton_info.get("klasa") or ""

    row += [
        float(L_beam),
        float(b_rect),
        float(h_rect),
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

        # dolne zbrojenie (z_*)
        _safe_float(st.session_state.get("z_ot_dolna", 5.0)),
        _safe_float(st.session_state.get("z_ot_gorna", 5.0)),
        _safe_float(st.session_state.get("z_ot_boczna", 5.0)),
        _safe_float(st.session_state.get("z_odst_poziomy", 5.0)),
        _safe_float(st.session_state.get("z_odst_pionowy", 5.0)),
        int(st.session_state.get("z_n_wlasne", 0)),
        int(st.session_state.get("z_warstwy_wlasne", 0)),

        # górne zbrojenie (s_*)
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
        beam_name_in = st.text_input(
            "Nazwa belki",
            value=str(st.session_state.get("beam_name_to_save", "")),
            key="beam_name_to_save",
        )

    confirm_overwrite_beam = st.checkbox(
        "Nadpisz istniejącą belkę o tej nazwie, jeśli istnieje",
        value=bool(st.session_state.get("chk_overwrite_beam", False)),
        key="chk_overwrite_beam",
    )

    with colL:
        submit_beam = st.form_submit_button("💾 Zapisz belkę", disabled=disabled_gs)

    if disabled_gs:
        st.info("Brak konfiguracji Google Sheets (SPREADSHEET_ID / gcp_service_account).")

# ✅ zapis poza st.form
if submit_beam:
    if not beam_name_in.strip():
        st.error("Podaj nazwę belki.")
    else:
        try:
            gc = gspread.authorize(CREDS)
            ss = gc.open_by_key(SPREADSHEET_ID)
            ws_beams = _get_or_create_worksheet(ss, SHEET_BEAMS_I, BEAM_HEADERS)

            # ✅ zawsze upewnij się, że nagłówki są poprawne
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

                    st.success(f"Nadpisano belkę „{beam_name_in}” w arkuszu „{SHEET_BEAMS_I}”. ID = {new_id}")
                else:
                    new_id = _next_id_from_df(existing_df)
                    new_row_df.loc[0, "ID"] = new_id

                    try:
                        ws_beams.append_rows(new_row_df.values.tolist(), value_input_option="RAW")
                    except Exception:
                        final_df = pd.concat([existing_df, new_row_df], ignore_index=True)
                        _update_sheet_atomic(ws_beams, BEAM_HEADERS, final_df)

                    st.success(f"Zapisano belkę „{beam_name_in}” do arkusza „{SHEET_BEAMS_I}”. ID = {new_id}")

        except Exception as e:
            st.error(f"Nie udało się zapisać belki: {e}")
