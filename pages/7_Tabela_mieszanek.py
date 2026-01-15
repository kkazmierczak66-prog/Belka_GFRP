from __future__ import annotations

import math
from typing import List
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# ---------- Konfiguracja strony ----------
st.set_page_config(page_title="Tabela mieszanek", page_icon="", layout="wide")
st.title("Tabela mieszanek")

# ---------- Google Sheets / Secrets ----------
GS_READY = False
try:
    from google.oauth2.service_account import Credentials
    import gspread
    from gspread.exceptions import WorksheetNotFound
    try:
        from gspread_dataframe import get_as_dataframe
    except Exception:
        get_as_dataframe = None

    GSA = "gcp_service_account"
    SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")
    SHEET_MATERIALS = st.secrets.get("SHEET_MATERIALS", "materiały")
    SHEET_RECIPES = st.secrets.get("SHEET_RECIPES", "receptury")
    SHEET_EXECUTIONS = st.secrets.get("SHEET_EXECUTIONS", "wykonania")
    SHEET_TESTS = st.secrets.get("SHEET_TESTS", "testy")

    if GSA in st.secrets and SPREADSHEET_ID:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        CREDS = Credentials.from_service_account_info(st.secrets[GSA], scopes=SCOPES)
        GS_READY = True
except Exception:
    GS_READY = False

# ====== KLUCZOWE: nagłówki jak w eksporcie z Receptury ======
HEADERS_RECIPES = [
    "timestamp", "recipe_name",
    "material_id", "nazwa", "kategoria",
    "gestosc_kgm3", "udzial_pct", "obj_m3", "masa_kgm3",
    "sum_obj_m3m3", "sum_mas_kgm3", "gestosc_mix_kgm3", "w_c",
    "fck_mpa", "fctm_mpa", "ecm_gpa"
]

EXEC_HEADER = ["recipe_name", "timestamp", "Nr wyk.", "Data wyk.", "Wykonawca/y", "Uwagi"]

# Nowy układ testów (zapis do Sheets):
TEST_HEADER = [
    "recipe_name",
    "timestamp",
    "Nr wyk.",
    "Nr testu",
    "Data testu",
    "Wiek próbki [dni]",
    "Rodzaj",
    "Masa próbki [g]",
    "Gęstość [kg/m3]",
    "Siła niszcząca [kN]",
    "Wynik [MPa]",
    "Wykonawca/y",
    "Opis zniszczenia / Uwagi",
]

# ---------- Pomocnicze: parser liczb 'tekstem' ----------
def to_num_pl(x):
    """
    Konwersja stringów typu "'123,45" lub "1 234,5" -> 123.45.
    Dla innych typów używa pd.to_numeric(errors='coerce').
    """
    if isinstance(x, str):
        s = x.strip()
        if s.startswith("'"):
            s = s[1:]
        s = s.replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return pd.NA
    return pd.to_numeric(x, errors="coerce")

def strip_apostrophe(s):
    if pd.isna(s):
        return s
    s = str(s)
    return s[1:] if s.startswith("'") else s

# ---------- NOWE: sanitizacja do Google Sheets (usuwa NaN/NA/inf) ----------
def gs_cell(x):
    """Zamienia wartości nie-JSON (NaN/NA/inf) na '' i normalizuje typy pod Sheets."""
    if x is None:
        return ""

    # pandas/numpy NA/NaN
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass

    # float nan/inf
    if isinstance(x, (float, np.floating)):
        xf = float(x)
        if math.isnan(xf) or math.isinf(xf):
            return ""
        return xf

    # numpy integers
    if isinstance(x, (np.integer,)):
        return int(x)

    # pandas Timestamp
    if isinstance(x, pd.Timestamp):
        if pd.isna(x):
            return ""
        return x.strftime("%d-%m-%Y")

    return x

# ---------- Geometria próbek ----------
RODZAJE = ["Kostka 10 cm", "Kostka 15 cm", "Cylinder 10 x 20 cm"]

# objętość [m3], pole przekroju [mm2]
SPECIMEN_GEOM = {
    "Kostka 10 cm": {
        "vol_m3": 0.1 * 0.1 * 0.1,                  # 0.001
        "area_mm2": (0.1 * 0.1) * 1_000_000.0,      # 0.01 m2 -> 10000 mm2
    },
    "Kostka 15 cm": {
        "vol_m3": 0.15 * 0.15 * 0.15,               # 0.003375
        "area_mm2": (0.15 * 0.15) * 1_000_000.0,    # 0.0225 m2 -> 22500 mm2
    },
    "Cylinder 10 x 20 cm": {
        "vol_m3": math.pi * (0.05 ** 2) * 0.20,     # pi*r^2*h
        "area_mm2": (math.pi * (0.05 ** 2)) * 1_000_000.0,  # pi*r^2 m2 -> mm2
    },
}

def compute_density_kgm3(mass_g, rodzaj: str):
    m = to_num_pl(mass_g)
    if pd.isna(m):
        return pd.NA
    if rodzaj not in SPECIMEN_GEOM:
        return pd.NA
    vol = SPECIMEN_GEOM[rodzaj]["vol_m3"]
    if not vol or vol <= 0:
        return pd.NA
    mass_kg = float(m) / 1000.0
    return mass_kg / vol

def compute_strength_mpa(force_kn, rodzaj: str):
    f = to_num_pl(force_kn)
    if pd.isna(f):
        return pd.NA
    if rodzaj not in SPECIMEN_GEOM:
        return pd.NA
    area_mm2 = SPECIMEN_GEOM[rodzaj]["area_mm2"]
    if not area_mm2 or area_mm2 <= 0:
        return pd.NA
    force_n = float(f) * 1000.0  # kN -> N
    return force_n / area_mm2     # N/mm2 == MPa

def normalize_tests_df(
    df: pd.DataFrame,
    base_exec_date: pd.Timestamp | None,
) -> pd.DataFrame:
    """
    Uzupełnia brakujące kolumny, czyści typy oraz liczy:
    - Wiek próbki [dni] = Data testu - Data wyk.
    - Gęstość [kg/m3] z masy i rodzaju
    - Wynik [MPa] z siły i rodzaju
    """
    cols = [
        "Nr testu",
        "Data testu",
        "Wiek próbki [dni]",
        "Rodzaj",
        "Masa próbki [g]",
        "Gęstość [kg/m3]",
        "Siła niszcząca [kN]",
        "Wynik [MPa]",
        "Wykonawca/y",
        "Opis zniszczenia / Uwagi",
    ]
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = pd.NA

    # Rodzaj: default gdy pusty
    out["Rodzaj"] = out["Rodzaj"].astype("object")
    out["Rodzaj"] = out["Rodzaj"].where(out["Rodzaj"].isin(RODZAJE), pd.NA)
    out["Rodzaj"] = out["Rodzaj"].fillna(RODZAJE[0])

    # Data testu: utrzymujemy jako tekst w edytorze, ale do obliczeń parsujemy
    dt_test = pd.to_datetime(out["Data testu"], errors="coerce", dayfirst=True)

    # Wiek próbki
    if base_exec_date is not None and not pd.isna(base_exec_date):
        age = (dt_test - base_exec_date).dt.days
    else:
        age = pd.Series([pd.NA] * len(out))
    out["Wiek próbki [dni]"] = age

    # Masa i siła na float do obliczeń
    mass_g = out["Masa próbki [g]"].apply(to_num_pl)
    force_kn = out["Siła niszcząca [kN]"].apply(to_num_pl)

    dens_list = []
    res_list = []
    for i in range(len(out)):
        rodz = out.loc[out.index[i], "Rodzaj"]
        dens_list.append(compute_density_kgm3(mass_g.iloc[i], rodz))
        res_list.append(compute_strength_mpa(force_kn.iloc[i], rodz))

    out["Gęstość [kg/m3]"] = pd.Series(dens_list, index=out.index)
    out["Wynik [MPa]"] = pd.Series(res_list, index=out.index)

    # Teksty – unikamy NaN w UI
    for c in ["Data testu", "Wykonawca/y", "Opis zniszczenia / Uwagi"]:
        out[c] = out[c].astype("object").where(pd.notna(out[c]), "")

    return out[cols]

# ---------- Obsługa Sheets ----------
def _open_ws(spreadsheet_id: str, sheet_name: str):
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    return ss.worksheet(sheet_name)

def _open_or_create_ws(spreadsheet_id: str, sheet_name: str, header: List[str]):
    """Otwiera arkusz, a jeśli nie istnieje – tworzy z podanym nagłówkiem."""
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    try:
        ws = ss.worksheet(sheet_name)
    except WorksheetNotFound:
        ws = ss.add_worksheet(title=sheet_name, rows=1000, cols=max(10, len(header)))
        ws.update("A1", [header])
    return ws

@st.cache_data(show_spinner=False)
def read_materials(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Czyta arkusz materiałów i normalizuje liczby na wypadek zapisu tekstowego."""
    ws = _open_ws(spreadsheet_id, sheet_name)
    if get_as_dataframe is not None:
        df = get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
    else:
        rows = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(rows)

    wanted = ["id", "nazwa", "kategoria", "gestosc_gcm3", "cena_pln", "co2e_kgkg"]
    for c in wanted:
        if c not in df.columns:
            df[c] = None

    df["id"] = df["id"].apply(strip_apostrophe)
    for c in ["id", "gestosc_gcm3", "cena_pln", "co2e_kgkg"]:
        df[c] = df[c].apply(to_num_pl)

    df["gestosc_kgm3"] = to_num_pl(df["gestosc_gcm3"]) * 1000.0
    return df[wanted + ["gestosc_kgm3"]]

@st.cache_data(show_spinner=False)
def read_recipes(spreadsheet_id: str, sheet_name: str, headers: List[str]) -> pd.DataFrame:
    """Czyta arkusz receptur zgodnie z HEADERS_RECIPES i liczy timestamp_dt."""
    ws = _open_ws(spreadsheet_id, sheet_name)
    vals = ws.get_all_values() or [headers]
    width = len(headers)
    fixed = []
    for row in vals:
        r = list(row)
        if len(r) < width:
            r += [""] * (width - len(r))
        else:
            r = r[:width]
        fixed.append(r)
    if not fixed:
        fixed = [headers]
    fixed[0] = headers
    df = pd.DataFrame(fixed[1:], columns=headers)

    if "material_id" in df.columns:
        df["material_id"] = df["material_id"].apply(strip_apostrophe)
        df["material_id"] = pd.to_numeric(df["material_id"], errors="coerce")

    num_cols = [
        "gestosc_kgm3", "udzial_pct", "obj_m3", "masa_kgm3",
        "sum_obj_m3m3", "sum_mas_kgm3", "gestosc_mix_kgm3", "w_c",
        "fck_mpa", "fctm_mpa", "ecm_gpa"
    ]
    for c in num_cols:
        if c in df.columns:
            df[c] = df[c].apply(to_num_pl)

    if "timestamp" in df.columns:
        def _to_dt_any(x):
            try:
                if isinstance(x, (int, float)) and not pd.isna(x):
                    return pd.to_datetime(x, unit="D", origin="1899-12-30", utc=True)
                return pd.to_datetime(str(x), utc=True, errors="coerce")
            except Exception:
                return pd.NaT
        df["timestamp_dt"] = df["timestamp"].apply(_to_dt_any)
    else:
        df["timestamp_dt"] = pd.NaT

    return df

@st.cache_data(show_spinner=False)
def read_executions_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    ws = _open_or_create_ws(spreadsheet_id, sheet_name, EXEC_HEADER)
    rows = ws.get_all_records(numericise_ignore=["all"])
    if not rows:
        return pd.DataFrame(columns=EXEC_HEADER)
    df = pd.DataFrame(rows)
    for c in EXEC_HEADER:
        if c not in df.columns:
            df[c] = None
    return df[EXEC_HEADER]

@st.cache_data(show_spinner=False)
def read_tests_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    ws = _open_or_create_ws(spreadsheet_id, sheet_name, TEST_HEADER)
    rows = ws.get_all_records(numericise_ignore=["all"])
    if not rows:
        return pd.DataFrame(columns=TEST_HEADER)

    df = pd.DataFrame(rows)

    # --- NOWE: normalizacja nazw kolumn (różnice typu m³ vs m3, spacje itp.) ---
    df.columns = [str(c).strip().replace("m³", "m3") for c in df.columns]

    # --- NOWE: dołóż brakujące kolumny z TEST_HEADER ---
    for c in TEST_HEADER:
        if c not in df.columns:
            df[c] = None

    # --- Zwracamy dokładnie w kolejności TEST_HEADER ---
    return df[TEST_HEADER]

def ensure_exec_ids(df_exec: pd.DataFrame) -> pd.DataFrame:
    if df_exec is None or df_exec.empty:
        return pd.DataFrame(columns=["Nr wyk.", "Data wyk.", "Wykonawca/y", "Uwagi"])
    out = df_exec.copy().reset_index(drop=True)
    out["Nr wyk."] = range(1, len(out) + 1)
    out["Nr wyk."] = pd.to_numeric(out["Nr wyk."], errors="coerce").astype("Int64")
    return out

def ensure_test_ids(df_tests: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "Nr testu",
        "Data testu",
        "Wiek próbki [dni]",
        "Rodzaj",
        "Masa próbki [g]",
        "Gęstość [kg/m3]",
        "Siła niszcząca [kN]",
        "Wynik [MPa]",
        "Wykonawca/y",
        "Opis zniszczenia / Uwagi",
    ]
    if df_tests is None or df_tests.empty:
        return pd.DataFrame(columns=cols)
    out = df_tests.copy().reset_index(drop=True)
    for c in cols:
        if c not in out.columns:
            out[c] = pd.NA
    out["Nr testu"] = range(1, len(out) + 1)
    out["Nr testu"] = pd.to_numeric(out["Nr testu"], errors="coerce").astype("Int64")
    return out[cols]

def stable_json_exec(df: pd.DataFrame) -> str:
    if df is None or len(df) == 0:
        return "[]"
    return (
        df.drop(columns=["_del"], errors="ignore")
        .sort_index()
        .sort_index(axis=1)
        .to_json(orient="records", date_format="iso", date_unit="s")
    )

def stable_json_test(df: pd.DataFrame) -> str:
    if df is None or len(df) == 0:
        return "[]"
    return (
        df.drop(columns=["_del"], errors="ignore")
        .sort_index()
        .sort_index(axis=1)
        .to_json(orient="records", date_format="iso", date_unit="s")
    )

def save_executions_and_tests_to_sheets():
    """Zapisuje wszystkie wykonania i testy z session_state do Google Sheets."""
    # --- wykonania ---
    exec_rows = []
    for key, value in st.session_state.items():
        if not (isinstance(key, str) and key.startswith("exec__") and key.endswith("__df")):
            continue
        ns_body = key[len("exec__"):-len("__df")]
        try:
            rname, ts_raw = ns_body.rsplit("__", 1)
        except ValueError:
            continue
        df_exec = ensure_exec_ids(value)
        for _, r in df_exec.iterrows():
            exec_rows.append({
                "recipe_name": rname,
                "timestamp": ts_raw,
                "Nr wyk.": r.get("Nr wyk.", ""),
                "Data wyk.": r.get("Data wyk.", ""),
                "Wykonawca/y": r.get("Wykonawca/y", ""),
                "Uwagi": r.get("Uwagi", ""),
            })

    ws_exec = _open_or_create_ws(SPREADSHEET_ID, SHEET_EXECUTIONS, EXEC_HEADER)
    ws_exec.clear()
    exec_values = [EXEC_HEADER]
    if exec_rows:
        exec_values += [[gs_cell(row.get(col, "")) for col in EXEC_HEADER] for row in exec_rows]
    ws_exec.update("A1", exec_values)

    # --- testy ---
    test_rows = []
    for key, value in st.session_state.items():
        if not (isinstance(key, str) and key.startswith("tests__") and key.endswith("__df")):
            continue
        ns_body = key[len("tests__"):-len("__df")]
        try:
            rname, ts_raw, nr_wyk_str = ns_body.rsplit("__", 2)
        except ValueError:
            continue
        try:
            nr_wyk_int = int(nr_wyk_str)
        except Exception:
            nr_wyk_int = None

        df_tests = ensure_test_ids(value)

        # zapisujemy również pola liczone (wiek, gęstość, wynik),
        # bo o to prosisz ("ma się zapisywać do google sheets")
        for _, r in df_tests.iterrows():
            test_rows.append({
                "recipe_name": rname,
                "timestamp": ts_raw,
                "Nr wyk.": nr_wyk_int,
                "Nr testu": r.get("Nr testu", ""),
                "Data testu": r.get("Data testu", ""),
                "Wiek próbki [dni]": r.get("Wiek próbki [dni]", ""),
                "Rodzaj": r.get("Rodzaj", ""),
                "Masa próbki [g]": r.get("Masa próbki [g]", ""),
                "Gęstość [kg/m3]": r.get("Gęstość [kg/m3]", ""),
                "Siła niszcząca [kN]": r.get("Siła niszcząca [kN]", ""),
                "Wynik [MPa]": r.get("Wynik [MPa]", ""),
                "Wykonawca/y": r.get("Wykonawca/y", ""),
                "Opis zniszczenia / Uwagi": r.get("Opis zniszczenia / Uwagi", ""),
            })

    ws_test = _open_or_create_ws(SPREADSHEET_ID, SHEET_TESTS, TEST_HEADER)
    ws_test.clear()
    test_values = [TEST_HEADER]
    if test_rows:
        test_values += [[gs_cell(row.get(col, "")) for col in TEST_HEADER] for row in test_rows]
    ws_test.update("A1", test_values)

def load_exec_state_from_sheet(rname: str, ts_raw: str, exec_ns: str):
    """Lazy load wykonań dla (recipe_name, timestamp)."""
    EXEC_DF = f"{exec_ns}__df"
    EXEC_SNAP = f"{exec_ns}__snap"
    EXEC_HIST = f"{exec_ns}__hist"

    if EXEC_DF in st.session_state:
        return

    try:
        df_all = read_executions_sheet(SPREADSHEET_ID, SHEET_EXECUTIONS)
        if not df_all.empty:
            mask = (df_all["recipe_name"].astype(str) == str(rname)) & (df_all["timestamp"].astype(str) == str(ts_raw))
            grp = df_all.loc[mask]
        else:
            grp = pd.DataFrame()
    except Exception:
        grp = pd.DataFrame()

    if grp is not None and not grp.empty:
        local = grp[["Nr wyk.", "Data wyk.", "Wykonawca/y", "Uwagi"]].copy()
        local = ensure_exec_ids(local)
    else:
        today = pd.Timestamp.now(tz=ZoneInfo("Europe/Warsaw")).date()
        local = pd.DataFrame([{
            "Nr wyk.": 1,
            "Data wyk.": today.strftime("%d-%m-%Y"),
            "Wykonawca/y": "",
            "Uwagi": "",
        }])
        local = ensure_exec_ids(local)

    st.session_state[EXEC_DF] = local
    st.session_state[EXEC_SNAP] = stable_json_exec(local)
    st.session_state[EXEC_HIST] = []

def load_tests_state_from_sheet(rname: str, ts_raw: str, nr_wyk_int: int, tests_ns: str):
    """Lazy load testów dla (recipe_name, timestamp, Nr wyk.)."""
    TDF = f"{tests_ns}__df"
    TSNAP = f"{tests_ns}__snap"

    if TDF in st.session_state:
        return

    try:
        df_all = read_tests_sheet(SPREADSHEET_ID, SHEET_TESTS)
        if not df_all.empty:
            mask = (
                (df_all["recipe_name"].astype(str) == str(rname))
                & (df_all["timestamp"].astype(str) == str(ts_raw))
                & (df_all["Nr wyk."].astype(str) == str(nr_wyk_int))
            )
            grp = df_all.loc[mask]
        else:
            grp = pd.DataFrame()
    except Exception:
        grp = pd.DataFrame()

    wanted_local_cols = [
        "Nr testu",
        "Data testu",
        "Wiek próbki [dni]",
        "Rodzaj",
        "Masa próbki [g]",
        "Gęstość [kg/m3]",
        "Siła niszcząca [kN]",
        "Wynik [MPa]",
        "Wykonawca/y",
        "Opis zniszczenia / Uwagi",
    ]

    if grp is not None and not grp.empty:
        local = grp.copy()
        # normalizacja nazw (strip + m³ -> m3), żeby uniknąć różnic w nagłówkach
        local.columns = [str(c).strip().replace("m³", "m3") for c in local.columns]

        for c in wanted_local_cols:
            if c not in local.columns:
                local[c] = pd.NA

        local = local[wanted_local_cols].copy()
        local = ensure_test_ids(local)
    else:
        local = pd.DataFrame([{
            "Nr testu": 1,
            "Data testu": "",
            "Wiek próbki [dni]": pd.NA,
            "Rodzaj": RODZAJE[0],
            "Masa próbki [g]": pd.NA,
            "Gęstość [kg/m3]": pd.NA,
            "Siła niszcząca [kN]": pd.NA,
            "Wynik [MPa]": pd.NA,
            "Wykonawca/y": "",
            "Opis zniszczenia / Uwagi": "",
        }])
        local = ensure_test_ids(local)

    st.session_state[TDF] = local
    st.session_state[TSNAP] = stable_json_test(local)

# ---------- Sprawdzenie konfiguracji ----------
if not GS_READY:
    st.error(
        "Brak konfiguracji Google Sheets. Uzupełnij SPREADSHEET_ID, SHEET_MATERIALS, "
        "SHEET_RECIPES, SHEET_EXECUTIONS, SHEET_TESTS oraz blok [gcp_service_account] "
        "w `.streamlit/secrets.toml`."
    )
    st.stop()

# ---------- Inicjalne wczytanie materiałów/receptur ----------
if "df_mat" not in st.session_state or "df_rec" not in st.session_state:
    try:
        st.session_state["df_mat"] = read_materials(SPREADSHEET_ID, SHEET_MATERIALS)
        st.session_state["df_rec"] = read_recipes(SPREADSHEET_ID, SHEET_RECIPES, HEADERS_RECIPES)
        st.toast("Wczytano dane z Google Sheets przy starcie ✔️", icon="✅")
    except Exception as e:
        st.error(f"Nie udało się wczytać danych z Google Sheets przy starcie: {e}")
        st.stop()

# ---------- Przyciski: Wczytaj / Zapisz ----------
btn_load_col, btn_save_col, _ = st.columns([1, 1, 4])

with btn_load_col:
    if st.button("↻ Odśwież", use_container_width=True):
        try:
            st.cache_data.clear()
            st.session_state["df_mat"] = read_materials(SPREADSHEET_ID, SHEET_MATERIALS)
            st.session_state["df_rec"] = read_recipes(SPREADSHEET_ID, SHEET_RECIPES, HEADERS_RECIPES)
            for key in list(st.session_state.keys()):
                if isinstance(key, str) and (key.startswith("exec__") or key.startswith("tests__")):
                    del st.session_state[key]
            st.toast("Wczytano dane z Google Sheets 🔄", icon="🔄")
            st.rerun()
        except Exception as e:
            st.error(f"Nie udało się ponownie wczytać danych: {e}")

with btn_save_col:
    if st.button("💾 Zapisz wszystko", use_container_width=True):
        try:
            save_executions_and_tests_to_sheets()
            st.toast("Zapisano wykonania i testy do Google Sheets ✔️", icon="✅")
        except Exception as e:
            st.error(f"Nie udało się zapisać danych wykonania/testów: {e}")

# ---------- Główne dane z session_state ----------
df_mat = st.session_state["df_mat"]
df_rec = st.session_state["df_rec"]

if df_rec.empty:
    st.info("Brak zapisanych receptur w arkuszu.")
    st.stop()

# ---------- Ostatnie wersje receptur ----------
is_summary = df_rec["nazwa"].astype(str).eq("__SUMMARY__")
df_sum = df_rec[is_summary].copy()
df_cmp = df_rec[~is_summary].copy()

if df_sum.empty:
    st.info("Nie znaleziono wierszy __SUMMARY__ w arkuszu receptur.")
    st.stop()

df_sum_sorted = df_sum.sort_values(["recipe_name", "timestamp_dt"], ascending=True)
latest_rows = df_sum_sorted.groupby("recipe_name", as_index=False).tail(1).copy()

latest_keys_df = latest_rows[["recipe_name", "timestamp"]].drop_duplicates()
df_cmp_latest = df_cmp.merge(latest_keys_df, on=["recipe_name", "timestamp"], how="inner")

# ---------- Obliczenia: w/s ----------
def _sum_mass(df: pd.DataFrame, cats: List[str]) -> pd.Series:
    mask = df["kategoria"].astype(str).str.lower().isin([c.lower() for c in cats])
    return df.loc[mask].groupby("recipe_name")["masa_kgm3"].sum(min_count=1)

mass_water = _sum_mass(df_cmp_latest, ["woda", "water"])
mass_binder = _sum_mass(df_cmp_latest, ["spoiwo", "cement", "binder"])
ws_series = (mass_water / mass_binder).replace([math.inf, -math.inf], pd.NA)

# ---------- Koszt i CO₂ ----------
df_cmp_latest["material_id"] = df_cmp_latest["material_id"].apply(strip_apostrophe)
df_cmp_latest["material_id"] = pd.to_numeric(df_cmp_latest["material_id"], errors="coerce")

df_mat_id = df_mat.rename(columns={"id": "material_id"})[["material_id", "cena_pln", "co2e_kgkg"]]
df_cost = df_cmp_latest.merge(df_mat_id, on="material_id", how="left")

df_cost["masa_kgm3"] = to_num_pl(df_cost["masa_kgm3"]).fillna(0.0)
df_cost["cena_pln"] = to_num_pl(df_cost["cena_pln"])
df_cost["co2e_kgkg"] = to_num_pl(df_cost["co2e_kgkg"])

df_cost["koszt_PLN_m3"] = df_cost["masa_kgm3"] * df_cost["cena_pln"]
df_cost["co2e_kg_m3"] = df_cost["masa_kgm3"] * df_cost["co2e_kgkg"]

agg_cost = df_cost.groupby("recipe_name")[["koszt_PLN_m3", "co2e_kg_m3"]].sum(min_count=1).reset_index()

# ---------- Tabela główna ----------
tbl = latest_rows[[
    "timestamp", "timestamp_dt", "recipe_name", "fck_mpa",
    "sum_obj_m3m3", "sum_mas_kgm3", "gestosc_mix_kgm3", "w_c"
]].copy()

tbl = tbl.merge(ws_series.rename("w_s").reset_index(), on="recipe_name", how="left")
tbl = tbl.merge(agg_cost, on="recipe_name", how="left")

for c in ["fck_mpa", "w_c", "w_s", "gestosc_mix_kgm3", "koszt_PLN_m3", "co2e_kg_m3"]:
    if c in tbl.columns:
        tbl[c] = to_num_pl(tbl[c])

tbl = tbl.sort_values("timestamp_dt", ascending=True)

flt = st.text_input("Filtr nazwy receptury (zawiera):", "")
if flt.strip():
    tbl = tbl[tbl["recipe_name"].str.contains(flt.strip(), case=False, na=False)]

if tbl.empty:
    st.info("Brak receptur po nałożeniu filtra.")
    st.stop()

tbl["Timestamp"] = pd.to_datetime(tbl["timestamp_dt"], utc=True).dt.tz_convert(
    ZoneInfo("Europe/Warsaw")
).dt.strftime("%d.%m.%Y %H:%M")

tbl["__key_recipe_name"] = tbl["recipe_name"].astype(str)
tbl["__key_timestamp"] = tbl["timestamp"].astype(str)

def _combine_fck(row):
    fck = row.get("fck_mpa", pd.NA)
    if pd.notna(fck):
        return f"{float(fck):.1f}"
    return ""

tbl["fck_teoretyczne_col"] = tbl.apply(_combine_fck, axis=1)

display_df = tbl.copy()
display_df["__select__"] = False

cols_display = [
    "__select__",
    "Timestamp",
    "recipe_name",
    "gestosc_mix_kgm3",
    "koszt_PLN_m3",
    "co2e_kg_m3",
    "fck_teoretyczne_col",
]

edited = st.data_editor(
    display_df[cols_display],
    use_container_width=True,
    hide_index=True,
    column_order=cols_display,
    column_config={
        "__select__": st.column_config.CheckboxColumn("", help="Zaznacz, aby usunąć lub zobaczyć szczegóły."),
        "Timestamp": "Timestamp",
        "recipe_name": "Nazwa",
        "gestosc_mix_kgm3": st.column_config.NumberColumn("Gęstość [kg/m³]", format="%.1f"),
        "koszt_PLN_m3": st.column_config.NumberColumn("Cena [PLN/m³]", format="%.2f"),
        "co2e_kg_m3": st.column_config.NumberColumn("GWP [CO₂/m³]", format="%.1f"),
        "fck_teoretyczne_col": st.column_config.TextColumn("fck, teoretyczne [MPa]"),
    },
    disabled=[
        "Timestamp", "recipe_name", "gestosc_mix_kgm3",
        "koszt_PLN_m3", "co2e_kg_m3", "fck_teoretyczne_col",
    ],
)

selected_idx = edited.index[edited["__select__"] == True] if "__select__" in edited.columns else []

def _delete_recipe_version(spreadsheet_id: str, sheet_name: str, recipe_name: str, timestamp_raw: str) -> int:
    ws = _open_ws(spreadsheet_id, sheet_name)
    values = ws.get_all_values()
    if not values:
        return 0
    headers = values[0]
    try:
        col_ts = headers.index("timestamp")
        col_name = headers.index("recipe_name")
    except ValueError:
        return 0

    to_delete = []
    for i, row in enumerate(values[1:], start=2):
        if len(row) > max(col_ts, col_name) and row[col_ts] == str(timestamp_raw) and row[col_name] == str(recipe_name):
            to_delete.append(i)
    for idx in sorted(to_delete, reverse=True):
        ws.delete_rows(idx)
    return len(to_delete)

del_col, _ = st.columns([1, 6])
with del_col:
    if st.button("🗑️ Usuń zaznaczone", type="primary", use_container_width=True):
        if len(selected_idx) == 0:
            st.info("Nie zaznaczono żadnych wierszy (użyj kolumny checkboxów po lewej).")
        else:
            keys = tbl.loc[selected_idx, ["__key_recipe_name", "__key_timestamp"]]
            total_deleted = 0
            for _, r in keys.iterrows():
                rn = str(r["__key_recipe_name"])
                ts_raw = str(r["__key_timestamp"])
                try:
                    total_deleted += _delete_recipe_version(SPREADSHEET_ID, SHEET_RECIPES, rn, ts_raw)
                except Exception as e:
                    st.error(f"Błąd usuwania {rn}: {e}")
            st.success(f"Usunięto {total_deleted} wierszy w arkuszu.")
            st.cache_data.clear()
            st.rerun()

# ---------- Wykonania + testy ----------
if len(selected_idx):
    st.subheader("Wykonania zaznaczonych receptur")
    sel_keys = tbl.loc[selected_idx, ["__key_recipe_name", "__key_timestamp", "Timestamp"]]

    for _, row in sel_keys.iterrows():
        rname = str(row["__key_recipe_name"])
        ts_raw = str(row["__key_timestamp"])
        ts_pretty = row["Timestamp"]

        with st.expander(f"{ts_pretty} • {rname}", expanded=False):
            st.markdown("## Wykonania")

            exec_ns = f"exec__{rname}__{ts_raw}"
            EXEC_DF = f"{exec_ns}__df"
            EXEC_SNAP = f"{exec_ns}__snap"
            EXEC_HIST = f"{exec_ns}__hist"
            EXEC_WKEY = f"{exec_ns}__editor"

            load_exec_state_from_sheet(rname, ts_raw, exec_ns)

            exec_df = ensure_exec_ids(st.session_state[EXEC_DF]).copy()
            if "Data wyk." in exec_df.columns:
                exec_df["Data wyk."] = (
                    exec_df["Data wyk."].astype("object").where(pd.notna(exec_df["Data wyk."]), "").astype(str)
                )

            exec_df["_del"] = False

            edited_exec = st.data_editor(
                exec_df,
                key=EXEC_WKEY,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                column_order=["_del", "Nr wyk.", "Data wyk.", "Wykonawca/y", "Uwagi"],
                column_config={
                    "_del": st.column_config.CheckboxColumn("Usuń?", help="Zaznacz wiersze do usunięcia"),
                    "Nr wyk.": st.column_config.NumberColumn("Nr wyk.", disabled=True),
                    "Data wyk.": st.column_config.TextColumn("Data wyk. (DD-MM-YYYY)"),
                    "Wykonawca/y": st.column_config.TextColumn("Wykonawca/y"),
                    "Uwagi": st.column_config.TextColumn("Uwagi"),
                },
            )

            c_add_exec, c_del_exec, _ = st.columns([1, 1, 4])

            with c_add_exec:
                if st.button("➕ Dodaj wykonanie", key=f"{exec_ns}__add"):
                    hist = st.session_state[EXEC_HIST]
                    hist.append(st.session_state[EXEC_DF].copy())
                    if len(hist) > 30:
                        hist.pop(0)
                    today = pd.Timestamp.now(tz=ZoneInfo("Europe/Warsaw")).date()
                    new_row = {
                        "Nr wyk.": len(st.session_state[EXEC_DF]) + 1,
                        "Data wyk.": today.strftime("%d-%m-%Y"),
                        "Wykonawca/y": "",
                        "Uwagi": "",
                    }
                    new_df = pd.concat([st.session_state[EXEC_DF], pd.DataFrame([new_row])], ignore_index=True)
                    new_df = ensure_exec_ids(new_df)
                    st.session_state[EXEC_DF] = new_df
                    st.session_state[EXEC_SNAP] = stable_json_exec(new_df)
                    st.rerun()

            with c_del_exec:
                if st.button("🗑️ Usuń zaznaczone wykonania", key=f"{exec_ns}__delbtn"):
                    if "_del" in edited_exec and bool(edited_exec["_del"].any()):
                        mask_keep = ~edited_exec["_del"].astype(bool)
                        new_exec = edited_exec.loc[mask_keep].drop(columns=["_del"])
                        new_exec = ensure_exec_ids(new_exec)

                        st.session_state[EXEC_HIST].append(st.session_state[EXEC_DF].copy())
                        if len(st.session_state[EXEC_HIST]) > 30:
                            st.session_state[EXEC_HIST].pop(0)

                        st.session_state[EXEC_DF] = new_exec
                        st.session_state[EXEC_SNAP] = stable_json_exec(new_exec)
                        st.toast(f"Usunięto {len(edited_exec) - len(new_exec)} wiersz(e) wykonania.", icon="🗑️")
                        st.rerun()
                    else:
                        st.info("Nie zaznaczono żadnych wykonań do usunięcia (użyj kolumny 'Usuń?').")

            edited_no_ui = edited_exec.drop(columns=["_del"], errors="ignore").copy()
            for col_name in ["Data wyk.", "Wykonawca/y", "Uwagi"]:
                if col_name in edited_no_ui.columns:
                    edited_no_ui[col_name] = edited_no_ui[col_name].astype("object").where(pd.notna(edited_no_ui[col_name]), "")
            edited_no_ui = ensure_exec_ids(edited_no_ui)

            current_exec = stable_json_exec(edited_no_ui)
            prev_exec_snap = st.session_state.get(EXEC_SNAP, "")
            if current_exec != prev_exec_snap:
                st.session_state[EXEC_HIST].append(st.session_state[EXEC_DF].copy())
                if len(st.session_state[EXEC_HIST]) > 30:
                    st.session_state[EXEC_HIST].pop(0)
                st.session_state[EXEC_DF] = edited_no_ui.copy()
                st.session_state[EXEC_SNAP] = current_exec
                st.toast("Zapisano zmiany w 'Wykonaniach' (lokalnie w sesji).", icon="✅")
                st.rerun()

            current_exec_df = ensure_exec_ids(st.session_state[EXEC_DF]).reset_index(drop=True)

            # === TESTY per wykonanie ===
            all_tests = []

            for idx_exec, erow in current_exec_df.iterrows():
                nr_wyk = erow.get("Nr wyk.", idx_exec + 1)
                try:
                    nr_wyk_int = int(nr_wyk)
                except Exception:
                    nr_wyk_int = idx_exec + 1

                base_date = pd.to_datetime(erow.get("Data wyk.", ""), errors="coerce", dayfirst=True)

                st.markdown(f"### Testy - Wykonanie {nr_wyk_int}")

                tests_ns = f"tests__{rname}__{ts_raw}__{nr_wyk_int}"
                TDF = f"{tests_ns}__df"
                TSNAP = f"{tests_ns}__snap"
                TWKEY = f"{tests_ns}__editor"

                load_tests_state_from_sheet(rname, ts_raw, nr_wyk_int, tests_ns)

                # Normalizacja + przeliczenia na start (żeby wiek/gęstość/wynik zawsze były)
                tests_df = ensure_test_ids(st.session_state[TDF].copy())
                tests_df = normalize_tests_df(tests_df, base_date)
                st.session_state[TDF] = tests_df.copy()

                tests_df["_del"] = False

                edited_tests = st.data_editor(
                    tests_df,
                    key=TWKEY,
                    num_rows="fixed",
                    use_container_width=True,
                    hide_index=True,
                    column_order=[
                        "_del",
                        "Nr testu",
                        "Data testu",
                        "Wiek próbki [dni]",
                        "Rodzaj",
                        "Masa próbki [g]",
                        "Gęstość [kg/m3]",
                        "Siła niszcząca [kN]",
                        "Wynik [MPa]",
                        "Wykonawca/y",
                        "Opis zniszczenia / Uwagi",
                    ],
                    column_config={
                        "_del": st.column_config.CheckboxColumn("Usuń?", help="Zaznacz testy do usunięcia"),
                        "Nr testu": st.column_config.NumberColumn("Nr testu", disabled=True),
                        "Data testu": st.column_config.TextColumn("Data testu (DD-MM-YYYY)"),
                        "Wiek próbki [dni]": st.column_config.NumberColumn("Wiek próbki [dni]", format="%.0f", disabled=True),
                        "Rodzaj": st.column_config.SelectboxColumn("Rodzaj", options=RODZAJE, default=RODZAJE[0]),
                        "Masa próbki [g]": st.column_config.NumberColumn("Masa próbki [g]", format="%.0f"),
                        "Gęstość [kg/m3]": st.column_config.NumberColumn("Gęstość [kg/m³]", format="%.1f", disabled=True),
                        "Siła niszcząca [kN]": st.column_config.NumberColumn("Siła niszcząca [kN]", format="%.2f"),
                        "Wynik [MPa]": st.column_config.NumberColumn("Wynik [MPa]", format="%.2f", disabled=True),
                        "Wykonawca/y": st.column_config.TextColumn("Wykonawca/y"),
                        "Opis zniszczenia / Uwagi": st.column_config.TextColumn("Opis zniszczenia / Uwagi"),
                    },
                    disabled=["Nr testu", "Wiek próbki [dni]", "Gęstość [kg/m3]", "Wynik [MPa]"],
                )

                c_add_test, c_del_test, _ = st.columns([1, 1, 4])

                with c_add_test:
                    if st.button("➕ Dodaj test", key=f"{tests_ns}__add"):
                        cur_tests = ensure_test_ids(st.session_state[TDF])
                        new_row_t = {
                            "Nr testu": None,
                            "Data testu": "",
                            "Wiek próbki [dni]": pd.NA,
                            "Rodzaj": RODZAJE[0],
                            "Masa próbki [g]": pd.NA,
                            "Gęstość [kg/m3]": pd.NA,
                            "Siła niszcząca [kN]": pd.NA,
                            "Wynik [MPa]": pd.NA,
                            "Wykonawca/y": "",
                            "Opis zniszczenia / Uwagi": "",
                        }
                        new_tests_df = pd.concat([cur_tests, pd.DataFrame([new_row_t])], ignore_index=True)
                        new_tests_df = ensure_test_ids(new_tests_df)
                        new_tests_df = normalize_tests_df(new_tests_df, base_date)
                        st.session_state[TDF] = new_tests_df
                        st.session_state[TSNAP] = stable_json_test(new_tests_df)
                        st.rerun()

                with c_del_test:
                    if st.button("🗑️ Usuń zaznaczone testy", key=f"{tests_ns}__del"):
                        if "_del" in edited_tests and bool(edited_tests["_del"].any()):
                            mask_keep_t = ~edited_tests["_del"].astype(bool)
                            new_tests = edited_tests.loc[mask_keep_t].drop(columns=["_del"])
                            new_tests = ensure_test_ids(new_tests)
                            new_tests = normalize_tests_df(new_tests, base_date)

                            st.session_state[TDF] = new_tests.copy()
                            st.session_state[TSNAP] = stable_json_test(new_tests)
                            st.toast(f"Usunięto {len(edited_tests) - len(new_tests)} test(y).", icon="🗑️")
                            st.rerun()
                        else:
                            st.info("Nie zaznaczono żadnych testów do usunięcia (użyj kolumny 'Usuń?').")

                # Zapis zmian (auto-przeliczenie) po edycji
                edited_no_ui = edited_tests.drop(columns=["_del"], errors="ignore").copy()
                edited_no_ui = ensure_test_ids(edited_no_ui)
                edited_no_ui = normalize_tests_df(edited_no_ui, base_date)

                cur_snap = stable_json_test(edited_no_ui)
                prev_snap = st.session_state.get(TSNAP, "")
                if cur_snap != prev_snap:
                    st.session_state[TDF] = edited_no_ui.copy()
                    st.session_state[TSNAP] = cur_snap
                    st.toast("Zapisano zmiany w 'Testach' (lokalnie w sesji).", icon="✅")
                    st.rerun()

                all_tests.append(edited_no_ui.copy())

            # --- Sekcja Wytrzymałość ---
            st.markdown("## Wytrzymałość")
            create_equiv = st.checkbox("Utwórz odpowiedniki", key=f"equiv__{rname}__{ts_raw}")

            if all_tests:
                tests_all_df = pd.concat(all_tests, ignore_index=True)

                tests_all_df["wiek_dni"] = pd.to_numeric(tests_all_df["Wiek próbki [dni]"], errors="coerce")
                tests_all_df["Wynik_num"] = pd.to_numeric(tests_all_df["Wynik [MPa]"], errors="coerce")

                df_scisk = tests_all_df.dropna(subset=["wiek_dni", "Wynik_num"]).copy()

                equiv_roz = pd.DataFrame()
                if create_equiv and not df_scisk.empty:
                    equiv_roz = df_scisk.copy()
                    equiv_roz["Wynik_num"] = 0.3 * (equiv_roz["Wynik_num"] ** (2 / 3))

                def fit_ec2_curve(days_arr, strength_arr):
                    t = np.clip(days_arr.astype(float), 1.0, None)
                    y = strength_arr.astype(float)
                    s_grid = np.linspace(0.15, 0.45, 61)
                    best = None
                    for s in s_grid:
                        beta = np.exp(s * (1 - np.sqrt(28.0 / t)))
                        denom = np.dot(beta, beta)
                        if denom <= 1e-12:
                            continue
                        f28 = np.dot(beta, y) / denom
                        y_hat = f28 * beta
                        err = np.mean((y_hat - y) ** 2)
                        if (best is None) or (err < best[0]):
                            best = (err, s, f28)
                    if best is None:
                        s = 0.25
                        f28 = np.nanmax(y) if len(y) else 0.0
                    else:
                        _, s, f28 = best
                    days_grid = np.arange(1, 29)
                    beta_grid = np.exp(s * (1 - np.sqrt(28.0 / days_grid)))
                    curve = f28 * beta_grid
                    return days_grid, curve

                def get_value_for_day(df_points, curve_days, curve_vals, target_day, tolerance):
                    exact = df_points[df_points["wiek_dni"] == target_day]
                    if not exact.empty:
                        return float(exact["Wynik_num"].iloc[0])

                    if tolerance > 0:
                        near = df_points[
                            (df_points["wiek_dni"] >= target_day - tolerance)
                            & (df_points["wiek_dni"] <= target_day + tolerance)
                        ]
                        if not near.empty:
                            idx = (near["wiek_dni"] - target_day).abs().idxmin()
                            return float(near.loc[idx, "Wynik_num"])

                    if target_day in curve_days:
                        val = float(curve_vals[curve_days == target_day])

                        past = df_points[df_points["wiek_dni"] < target_day]
                        future = df_points[df_points["wiek_dni"] > target_day]

                        if not past.empty:
                            past_idx = (past["wiek_dni"] - target_day).abs().idxmin()
                            past_val = float(past.loc[past_idx, "Wynik_num"])
                            if val < past_val:
                                return past_val

                        if not future.empty:
                            fut_idx = (future["wiek_dni"] - target_day).abs().idxmin()
                            fut_val = float(future.loc[fut_idx, "Wynik_num"])
                            if val > fut_val:
                                return fut_val

                        return val

                    return None

                def build_table(df_points, curve_days, curve_vals):
                    params = [("fck,3", 3, 0), ("fck,7", 7, 1), ("fck,14", 14, 1), ("fck,28", 28, 2)]
                    rows = {}
                    for label, day, tol in params:
                        rows[label] = get_value_for_day(df_points, curve_days, curve_vals, day, tol)

                    # monotoniczność
                    if rows["fck,7"] is not None and rows["fck,3"] is not None and rows["fck,7"] < rows["fck,3"]:
                        rows["fck,7"] = rows["fck,3"]
                    if rows["fck,14"] is not None and rows["fck,7"] is not None and rows["fck,14"] < rows["fck,7"]:
                        rows["fck,14"] = rows["fck,7"]
                    if rows["fck,28"] is not None and rows["fck,14"] is not None and rows["fck,28"] < rows["fck,14"]:
                        rows["fck,28"] = rows["fck,14"]

                    return pd.DataFrame([{"Parametr": k, "Wartość [MPa]": (v if v is not None else "-")} for k, v in rows.items()])

                col1, col2 = st.columns(2)

                with col1:
                    if not df_scisk.empty:
                        fig1, ax1 = plt.subplots(figsize=(5, 3))
                        ax1.plot(df_scisk["wiek_dni"], df_scisk["Wynik_num"], "o", color="#1f77b4")
                        days_c, curve_c = fit_ec2_curve(df_scisk["wiek_dni"].values, df_scisk["Wynik_num"].values)
                        ax1.plot(days_c, curve_c, "-", color="#2ca02c", linewidth=1.5)
                        ax1.set_xlabel("Wiek próbki [dni]")
                        ax1.set_ylabel("Wytrzymałość [MPa]")
                        ax1.set_title("Ściskanie — punkty + EC2 (dopas.)")
                        ax1.grid(True)
                        max_day = max(int(np.nanmax([df_scisk["wiek_dni"].max()])), 28)
                        ax1.set_xlim(0, max_day)
                        ax1.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
                        ax1.set_ylim(bottom=0)
                        fig1.tight_layout()
                        st.pyplot(fig1, use_container_width=False)

                with col2:
                    if create_equiv and not equiv_roz.empty:
                        fig2, ax2 = plt.subplots(figsize=(5, 3))
                        ax2.plot(equiv_roz["wiek_dni"], equiv_roz["Wynik_num"], "x", color="#7f7f7f")
                        days_t, curve_t = fit_ec2_curve(equiv_roz["wiek_dni"].values, equiv_roz["Wynik_num"].values)
                        ax2.plot(days_t, curve_t, "-", color="#d62728", linewidth=1.5)
                        ax2.set_xlabel("Wiek próbki [dni]")
                        ax2.set_ylabel("Wytrzymałość [MPa]")
                        ax2.set_title("Rozciąganie (odpow.) — punkty + EC2 (dopas.)")
                        ax2.grid(True)
                        max_day = max(int(np.nanmax([equiv_roz["wiek_dni"].max()])), 28)
                        ax2.set_xlim(0, max_day)
                        ax2.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
                        ax2.set_ylim(bottom=0)
                        fig2.tight_layout()
                        st.pyplot(fig2, use_container_width=False)

                col1t, col2t = st.columns(2)
                with col1t:
                    if not df_scisk.empty:
                        days_c2, curve_c2 = fit_ec2_curve(df_scisk["wiek_dni"].values, df_scisk["Wynik_num"].values)
                        table_scisk = build_table(df_scisk, days_c2, curve_c2)
                        st.subheader("Tabela Ściskanie")
                        st.dataframe(table_scisk, use_container_width=True)

                with col2t:
                    if create_equiv and not equiv_roz.empty:
                        days_t2, curve_t2 = fit_ec2_curve(equiv_roz["wiek_dni"].values, equiv_roz["Wynik_num"].values)
                        table_roz = build_table(equiv_roz, days_t2, curve_t2)
                        st.subheader("Tabela Rozciąganie (odpow.)")
                        st.dataframe(table_roz, use_container_width=True)
