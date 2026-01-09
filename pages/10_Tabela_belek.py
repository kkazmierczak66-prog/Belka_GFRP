from __future__ import annotations

import math
from typing import Any, List
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

# ============================================================
# Konfiguracja strony
# ============================================================
st.set_page_config(page_title="Tabela belek", page_icon="", layout="wide")
st.title("Tabela belek")

# ============================================================
# Google Sheets / Secrets
# ============================================================
GS_READY = False
try:
    from google.oauth2.service_account import Credentials
    import gspread
    from gspread.exceptions import WorksheetNotFound

    GSA = "gcp_service_account"
    SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")

    # arkusze z belkami
    SHEET_BEAMS_I = st.secrets.get("SHEET_BEAMS_I", "belki i")
    SHEET_BEAMS_TPD = st.secrets.get("SHEET_BEAMS_TPD", "belki tpd")

    # arkusze: wykonania/testy belek (logika expanderów jak w Twojej "tabela mieszanek")
    SHEET_BEAM_EXECUTIONS = st.secrets.get("SHEET_BEAM_EXECUTIONS", "wykonania_belek")
    SHEET_BEAM_TESTS = st.secrets.get("SHEET_BEAM_TESTS", "testy_belek")

    if GSA in st.secrets and SPREADSHEET_ID:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        CREDS = Credentials.from_service_account_info(st.secrets[GSA], scopes=SCOPES)
        GS_READY = True
except Exception:
    GS_READY = False

if not GS_READY:
    st.error(
        "Brak konfiguracji Google Sheets. Uzupełnij SPREADSHEET_ID oraz blok [gcp_service_account] w secrets.toml."
    )
    st.stop()

# ============================================================
# Stałe: nagłówki wykonania / testy (dla belek)
# ============================================================
BEAM_EXEC_HEADER = ["beam_key", "Nr wyk.", "Data wyk.", "Wykonawca/y", "Uwagi"]
BEAM_TEST_HEADER = ["beam_key", "Nr wyk.", "Nr testu", "Data testu", "Wynik", "Wykonawca/y", "Uwagi"]

# ============================================================
# Pomocnicze
# ============================================================
def to_num_pl(x):
    """Konwersja stringów typu "'123,45" / "1 234,5" -> 123.45; inne -> pd.to_numeric(coerce)."""
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

def _canon_geom(x: Any) -> str:
    s = str(x or "").strip().lower()
    s = s.replace(" ", "")
    # ujednolicenia
    if s in {"i", "prost", "prost.", "prosta", "prostokąt", "prostokat"}:
        return "i"
    if s in {"tpd", "t-pd", "t_pd"}:
        return "tpd"
    return s or "?"

def _normalize_name(s: str) -> str:
    return " ".join(str(s or "").split()).strip().lower()


def _open_ws(spreadsheet_id: str, sheet_name: str):
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    return ss.worksheet(sheet_name)


def _open_or_create_ws(spreadsheet_id: str, sheet_name: str, header: List[str]):
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(spreadsheet_id)
    try:
        ws = ss.worksheet(sheet_name)
    except WorksheetNotFound:
        ws = ss.add_worksheet(title=sheet_name, rows=2000, cols=max(10, len(header)))
        ws.update("A1", [header])
    return ws


def _safe_get_col(df: pd.DataFrame, candidates: List[str], default=None):
    for c in candidates:
        if c in df.columns:
            return df[c]
    return default


def _sheet_to_df(ws, fallback_headers: List[str] | None = None) -> pd.DataFrame:
    vals = ws.get_all_values()
    if not vals:
        if fallback_headers:
            return pd.DataFrame(columns=fallback_headers)
        return pd.DataFrame()

    header = vals[0]
    df = pd.DataFrame(vals[1:], columns=header)

    # jeśli arkusz jest pusty ale ma header
    if df.empty:
        return pd.DataFrame(columns=header)

    return df


def ensure_exec_ids(df_exec: pd.DataFrame) -> pd.DataFrame:
    if df_exec is None or df_exec.empty:
        return pd.DataFrame(columns=["Nr wyk.", "Data wyk.", "Wykonawca/y", "Uwagi"])
    out = df_exec.copy().reset_index(drop=True)
    out["Nr wyk."] = range(1, len(out) + 1)
    out["Nr wyk."] = pd.to_numeric(out["Nr wyk."], errors="coerce").astype("Int64")
    return out


def ensure_test_ids(df_tests: pd.DataFrame) -> pd.DataFrame:
    if df_tests is None or df_tests.empty:
        return pd.DataFrame(columns=["Nr testu", "Data testu", "Wynik", "Wykonawca/y", "Uwagi"])
    out = df_tests.copy().reset_index(drop=True)
    out["Nr testu"] = range(1, len(out) + 1)
    out["Nr testu"] = pd.to_numeric(out["Nr testu"], errors="coerce").astype("Int64")
    return out


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


# ============================================================
# Czytanie arkuszy belek
# ============================================================
@st.cache_data(show_spinner=False)
def read_beams_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    ws = _open_ws(spreadsheet_id, sheet_name)
    df = _sheet_to_df(ws)
    if df is None or df.empty:
        return pd.DataFrame()
    return df


@st.cache_data(show_spinner=False)
def read_beam_exec_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    ws = _open_or_create_ws(spreadsheet_id, sheet_name, BEAM_EXEC_HEADER)
    rows = ws.get_all_records(numericise_ignore=["all"])
    if not rows:
        return pd.DataFrame(columns=BEAM_EXEC_HEADER)
    df = pd.DataFrame(rows)
    for c in BEAM_EXEC_HEADER:
        if c not in df.columns:
            df[c] = None
    return df[BEAM_EXEC_HEADER]


@st.cache_data(show_spinner=False)
def read_beam_tests_sheet(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    ws = _open_or_create_ws(spreadsheet_id, sheet_name, BEAM_TEST_HEADER)
    rows = ws.get_all_records(numericise_ignore=["all"])
    if not rows:
        return pd.DataFrame(columns=BEAM_TEST_HEADER)
    df = pd.DataFrame(rows)
    for c in BEAM_TEST_HEADER:
        if c not in df.columns:
            df[c] = None
    return df[BEAM_TEST_HEADER]


# ============================================================
# Zapis wykonania/testy do Sheets (jak w Twoim pliku mieszanek)
# ============================================================
def save_beam_exec_and_tests_to_sheets():
    # --- wykonania ---
    exec_rows = []
    for key, value in st.session_state.items():
        if not (isinstance(key, str) and key.startswith("exec__") and key.endswith("__df")):
            continue
        # exec__{beam_key}__df
        beam_key = key[len("exec__") : -len("__df")]
        df_exec = ensure_exec_ids(value)
        for _, r in df_exec.iterrows():
            exec_rows.append(
                {
                    "beam_key": beam_key,
                    "Nr wyk.": r.get("Nr wyk.", ""),
                    "Data wyk.": r.get("Data wyk.", ""),
                    "Wykonawca/y": r.get("Wykonawca/y", ""),
                    "Uwagi": r.get("Uwagi", ""),
                }
            )

    ws_exec = _open_or_create_ws(SPREADSHEET_ID, SHEET_BEAM_EXECUTIONS, BEAM_EXEC_HEADER)
    ws_exec.clear()
    exec_values = [BEAM_EXEC_HEADER] + [[row.get(col, "") for col in BEAM_EXEC_HEADER] for row in exec_rows] if exec_rows else [BEAM_EXEC_HEADER]
    ws_exec.update("A1", exec_values)

    # --- testy ---
    test_rows = []
    for key, value in st.session_state.items():
        if not (isinstance(key, str) and key.startswith("tests__") and key.endswith("__df")):
            continue
        # tests__{beam_key}__{nr_wyk}__df
        body = key[len("tests__") : -len("__df")]
        try:
            beam_key, nr_wyk_str = body.rsplit("__", 1)
        except ValueError:
            continue
        try:
            nr_wyk_int = int(nr_wyk_str)
        except Exception:
            nr_wyk_int = None

        df_tests = ensure_test_ids(value)
        for _, r in df_tests.iterrows():
            test_rows.append(
                {
                    "beam_key": beam_key,
                    "Nr wyk.": nr_wyk_int,
                    "Nr testu": r.get("Nr testu", ""),
                    "Data testu": r.get("Data testu", ""),
                    "Wynik": r.get("Wynik", ""),
                    "Wykonawca/y": r.get("Wykonawca/y", ""),
                    "Uwagi": r.get("Uwagi", ""),
                }
            )

    ws_test = _open_or_create_ws(SPREADSHEET_ID, SHEET_BEAM_TESTS, BEAM_TEST_HEADER)
    ws_test.clear()
    test_values = [BEAM_TEST_HEADER] + [[row.get(col, "") for col in BEAM_TEST_HEADER] for row in test_rows] if test_rows else [BEAM_TEST_HEADER]
    ws_test.update("A1", test_values)


# ============================================================
# Lazy load: wykonania/testy dla belki
# ============================================================
def load_exec_state_from_sheet(beam_key: str):
    EXEC_DF = f"exec__{beam_key}__df"
    EXEC_SNAP = f"exec__{beam_key}__snap"
    EXEC_HIST = f"exec__{beam_key}__hist"

    if EXEC_DF in st.session_state:
        return

    try:
        df_all = read_beam_exec_sheet(SPREADSHEET_ID, SHEET_BEAM_EXECUTIONS)
        grp = df_all[df_all["beam_key"].astype(str) == str(beam_key)]
    except Exception:
        grp = pd.DataFrame()

    if grp is not None and not grp.empty:
        local = grp[["Nr wyk.", "Data wyk.", "Wykonawca/y", "Uwagi"]].copy()
        local = ensure_exec_ids(local)
    else:
        today = pd.Timestamp.now(tz=ZoneInfo("Europe/Warsaw")).date()
        local = pd.DataFrame(
            [
                {
                    "Nr wyk.": 1,
                    "Data wyk.": today.strftime("%d-%m-%Y"),
                    "Wykonawca/y": "",
                    "Uwagi": "",
                }
            ]
        )
        local = ensure_exec_ids(local)

    st.session_state[EXEC_DF] = local
    st.session_state[EXEC_SNAP] = stable_json_exec(local)
    st.session_state[EXEC_HIST] = []


def load_tests_state_from_sheet(beam_key: str, nr_wyk_int: int):
    TDF = f"tests__{beam_key}__{nr_wyk_int}__df"
    TSNAP = f"tests__{beam_key}__{nr_wyk_int}__snap"

    if TDF in st.session_state:
        return

    try:
        df_all = read_beam_tests_sheet(SPREADSHEET_ID, SHEET_BEAM_TESTS)
        grp = df_all[
            (df_all["beam_key"].astype(str) == str(beam_key))
            & (df_all["Nr wyk."].astype(str) == str(nr_wyk_int))
        ]
    except Exception:
        grp = pd.DataFrame()

    if grp is not None and not grp.empty:
        local = grp[["Nr testu", "Data testu", "Wynik", "Wykonawca/y", "Uwagi"]].copy()
        local = ensure_test_ids(local)
    else:
        local = pd.DataFrame(
            [
                {
                    "Nr testu": 1,
                    "Data testu": "",
                    "Wynik": "",
                    "Wykonawca/y": "",
                    "Uwagi": "",
                }
            ]
        )
        local = ensure_test_ids(local)

    st.session_state[TDF] = local
    st.session_state[TSNAP] = stable_json_test(local)


# ============================================================
# Min z testów: arkusz + session_state (żeby działało od razu)
# ============================================================
def _collect_tests_from_session_state() -> pd.DataFrame:
    rows = []
    for key, value in st.session_state.items():
        if not (isinstance(key, str) and key.startswith("tests__") and key.endswith("__df")):
            continue
        body = key[len("tests__") : -len("__df")]
        try:
            beam_key, _nr_wyk_str = body.rsplit("__", 1)
        except ValueError:
            continue

        try:
            df = value.copy()
        except Exception:
            continue

        if df is None or df.empty or "Wynik" not in df.columns:
            continue

        for w in df["Wynik"].tolist():
            rows.append({"beam_key": beam_key, "Wynik": w})

    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["beam_key", "Wynik"])


@st.cache_data(show_spinner=False)
def _read_tests_min_from_sheet(spreadsheet_id: str, sheet_name: str) -> pd.Series:
    df = read_beam_tests_sheet(spreadsheet_id, sheet_name)
    if df is None or df.empty or "beam_key" not in df.columns or "Wynik" not in df.columns:
        return pd.Series(dtype="float64")

    tmp = df[["beam_key", "Wynik"]].copy()
    tmp["Wynik_num"] = tmp["Wynik"].apply(to_num_pl)
    tmp = tmp.dropna(subset=["Wynik_num"])
    if tmp.empty:
        return pd.Series(dtype="float64")
    return tmp.groupby("beam_key")["Wynik_num"].min()


def compute_test_min_per_beam() -> pd.Series:
    s_sheet = _read_tests_min_from_sheet(SPREADSHEET_ID, SHEET_BEAM_TESTS)

    df_ss = _collect_tests_from_session_state()
    if df_ss.empty:
        return s_sheet

    df_ss["Wynik_num"] = df_ss["Wynik"].apply(to_num_pl)
    df_ss = df_ss.dropna(subset=["Wynik_num"])
    if df_ss.empty:
        return s_sheet

    s_ss = df_ss.groupby("beam_key")["Wynik_num"].min()
    if s_sheet is None or s_sheet.empty:
        return s_ss

    return pd.concat([s_sheet, s_ss], axis=1).min(axis=1)


# ============================================================
# UI: Odśwież / Zapisz wszystko
# ============================================================
btn_load_col, btn_save_col, _ = st.columns([1, 1, 4])

with btn_load_col:
    if st.button("↻ Odśwież", use_container_width=True):
        st.cache_data.clear()
        # czyścimy lokalne wykonania/testy -> odtworzą się z arkusza przy otwarciu expanderów
        for k in list(st.session_state.keys()):
            if isinstance(k, str) and (k.startswith("exec__") or k.startswith("tests__")):
                del st.session_state[k]
        st.rerun()

with btn_save_col:
    if st.button("💾 Zapisz wszystko", use_container_width=True):
        try:
            save_beam_exec_and_tests_to_sheets()
            st.toast("Zapisano wykonania i testy", icon="✅")
            st.cache_data.clear()
        except Exception as e:
            st.error(f"Nie udało się zapisać wykonania/testów: {e}")



# ============================================================
# Wczytanie belek z dwóch arkuszy + scalenie (kolejność: belki i -> belki tpd)
# ============================================================
df_i = read_beams_sheet(SPREADSHEET_ID, SHEET_BEAMS_I)
df_tpd = read_beams_sheet(SPREADSHEET_ID, SHEET_BEAMS_TPD)

if df_i.empty and df_tpd.empty:
    st.info("Brak belek w arkuszach.")
    st.stop()

def _build_beams_view(df_src: pd.DataFrame, geom_label: str) -> pd.DataFrame:
    if df_src is None or df_src.empty:
        return pd.DataFrame()

    # --- mapowanie kolumn (elastyczne) ---
    col_id = "ID" if "ID" in df_src.columns else None
    col_name = "Nazwa belki" if "Nazwa belki" in df_src.columns else ("Nazwa" if "Nazwa" in df_src.columns else None)

    # u Ciebie w arkuszu bywa "Receptura beton" (bez "u")
    col_mix = None
    for cand in ["Receptura betonu", "Receptura beton", "Mieszanka"]:
        if cand in df_src.columns:
            col_mix = cand
            break

    # geometria z arkusza (jeśli jest)
    col_geom = "Geometria" if "Geometria" in df_src.columns else None

    # cena mieszanki [USD/m³]
    col_mix_price_m3 = None
    for cand in [
        "Cena mieszanki [USD/m³]",
        "Cena mieszanki [USD/m3]",
        "Cena mieszanki [USD/m^3]",
        "Cena mieszanki / m3 [USD]",
        "Cena mieszanki USD/m3",
    ]:
        if cand in df_src.columns:
            col_mix_price_m3 = cand
            break

    # procedury
    col_p_aci = "P_ACI_440_kN" if "P_ACI_440_kN" in df_src.columns else None
    col_p_jsce = "P_JSCE_kN" if "P_JSCE_kN" in df_src.columns else None
    col_p_csa = "P_CSA_kN" if "P_CSA_kN" in df_src.columns else None

    # Pmin/własne + wyniki
    col_p_min = "P_min_proc_kN" if "P_min_proc_kN" in df_src.columns else None
    col_p_custom = "P_custom_kN" if "P_custom_kN" in df_src.columns else None
    col_w_min = "Wynik_min_proc_USD_per_kN" if "Wynik_min_proc_USD_per_kN" in df_src.columns else None
    col_w_custom = "Wynik_custom_USD_per_kN" if "Wynik_custom_USD_per_kN" in df_src.columns else None

    # punktacja (kolumny “kosztowo-masowe”)
    punkt_cols = [
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
    ]

    # jeśli w źródle ich nie ma, to je tworzę (żeby potem nie było KeyError)
    for c in punkt_cols:
        if c not in df_src.columns:
            df_src[c] = ""

    out = pd.DataFrame()
    out["ID"] = df_src[col_id] if col_id else ""
    out["Nazwa"] = df_src[col_name] if col_name else ""
    out["Mieszanka"] = df_src[col_mix] if col_mix else ""
    out["Cena mieszanki [USD/m³]"] = df_src[col_mix_price_m3] if col_mix_price_m3 else ""

    # pokazuj tekst z arkusza (np. "prost."), ale kanonizuj do kluczy
    if col_geom:
        out["Geometria"] = df_src[col_geom].astype(str).fillna("").apply(lambda v: v.strip())
    else:
        out["Geometria"] = geom_label

    out["__geom_id"] = out["Geometria"].apply(_canon_geom)

    out["P_ACI_440 [kN]"] = df_src[col_p_aci] if col_p_aci else ""
    out["P_JSCE [kN]"] = df_src[col_p_jsce] if col_p_jsce else ""
    out["P_CSA [kN]"] = df_src[col_p_csa] if col_p_csa else ""

    for c in punkt_cols:
        out[c] = df_src[c]

    out["P,min [kN]"] = df_src[col_p_min] if col_p_min else ""
    out["P,własne [kN]"] = df_src[col_p_custom] if col_p_custom else ""
    out["Wynik,min [USD/kN]"] = df_src[col_w_min] if col_w_min else ""
    out["Wynik,własne [USD/kN]"] = df_src[col_w_custom] if col_w_custom else ""

    # stabilny beam_key: używaj __geom_id (żeby "prost." -> "i")
    def _mk_key(r):
        rid = str(r.get("ID", "")).strip()
        if rid:
            return f"{r.get('__geom_id', '?')}|{rid}"
        return f"{r.get('__geom_id', '?')}|{_normalize_name(r.get('Nazwa', ''))}"

    out["__beam_key"] = out.apply(_mk_key, axis=1)
    return out

    # procedury
    col_p_aci = "P_ACI_440_kN" if "P_ACI_440_kN" in df_src.columns else None
    col_p_jsce = "P_JSCE_kN" if "P_JSCE_kN" in df_src.columns else None
    col_p_csa = "P_CSA_kN" if "P_CSA_kN" in df_src.columns else None

    # Pmin/własne + wyniki
    col_p_min = "P_min_proc_kN" if "P_min_proc_kN" in df_src.columns else None
    col_p_custom = "P_custom_kN" if "P_custom_kN" in df_src.columns else None
    col_w_min = "Wynik_min_proc_USD_per_kN" if "Wynik_min_proc_USD_per_kN" in df_src.columns else None
    col_w_custom = "Wynik_custom_USD_per_kN" if "Wynik_custom_USD_per_kN" in df_src.columns else None

    # punktacja (jak w zapisach)
    punkt_cols = [
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
    ]
    for c in punkt_cols:
        if c not in df_src.columns:
            df_src[c] = ""

    out = pd.DataFrame()
    out["ID"] = df_src[col_id] if col_id else ""
    out["Nazwa"] = df_src[col_name] if col_name else ""
    out["Mieszanka"] = df_src[col_mix] if col_mix else ""
    out["Cena mieszanki [USD/m³]"] = df_src[col_mix_price_m3] if col_mix_price_m3 else ""
    out["Geometria"] = geom_label

    out["P_ACI_440 [kN]"] = df_src[col_p_aci] if col_p_aci else ""
    out["P_JSCE [kN]"] = df_src[col_p_jsce] if col_p_jsce else ""
    out["P_CSA [kN]"] = df_src[col_p_csa] if col_p_csa else ""

    for c in punkt_cols:
        out[c] = df_src[c]

    out["P,min [kN]"] = df_src[col_p_min] if col_p_min else ""
    out["P,własne [kN]"] = df_src[col_p_custom] if col_p_custom else ""
    out["Wynik,min [USD/kN]"] = df_src[col_w_min] if col_w_min else ""
    out["Wynik,własne [USD/kN]"] = df_src[col_w_custom] if col_w_custom else ""

    # beam_key (stabilny i unikalny): geometria|ID jeśli jest, inaczej geometria|nazwa_norm
    def _mk_key(r):
        rid = str(r.get("ID", "")).strip()
        if rid:
            return f"{geom_label}|{rid}"
        return f"{geom_label}|{_normalize_name(r.get('Nazwa', ''))}"

    out["__beam_key"] = out.apply(_mk_key, axis=1)
    return out


view_i = _build_beams_view(df_i, "i")
view_tpd = _build_beams_view(df_tpd, "tpd")

beams = pd.concat([view_i, view_tpd], ignore_index=True)
beams["__geom_order"] = beams["__geom_id"].map({"i": 0, "tpd": 1}).fillna(9).astype(int)


# liczby -> numeric gdzie sensownie
num_candidates = [
    "Cena mieszanki [USD/m³]",
    "P_ACI_440 [kN]", "P_JSCE [kN]", "P_CSA [kN]",
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
    "P,min [kN]", "P,własne [kN]",
    "Wynik,min [USD/kN]", "Wynik,własne [USD/kN]",
]
for c in num_candidates:
    if c in beams.columns:
        beams[c] = beams[c].apply(to_num_pl)
# === wylicz Cena mieszanki [USD/m³] jeśli brak / NaN ===
if "Cena mieszanki [USD/m³]" not in beams.columns:
    beams["Cena mieszanki [USD/m³]"] = pd.NA

# objętość w litrach -> m³
vol_l = beams["Łączna obj. belki [l]"].apply(to_num_pl)
price_per_beam = beams["Cena mieszanki / belkę [USD]"].apply(to_num_pl)

vol_m3 = vol_l / 1000.0

calc_mix_price_m3 = (price_per_beam / vol_m3.where(vol_m3 > 0)).replace([math.inf, -math.inf], pd.NA)

# nadpisz tylko tam, gdzie obecna wartość jest pusta/NaN
mask_missing = beams["Cena mieszanki [USD/m³]"].isna()
beams.loc[mask_missing, "Cena mieszanki [USD/m³]"] = calc_mix_price_m3.loc[mask_missing]


beams = beams.sort_values(["__geom_order", "Nazwa"], ascending=[True, True]).reset_index(drop=True)

# ============================================================
# Współczynnik raportu (edytowalny) – trzymamy w session_state mapą
# ============================================================
if "report_coef_map" not in st.session_state:
    st.session_state["report_coef_map"] = {}  # beam_key -> float

coef_map: dict = st.session_state["report_coef_map"]

def _get_coef(beam_key: str) -> float:
    v = coef_map.get(beam_key, None)
    try:
        v = float(v)
        if v <= 0:
            return 1.0
        return v
    except Exception:
        return 1.0

beams["Wsp. Raport"] = beams["__beam_key"].apply(_get_coef)

# ============================================================
# P,min(zbadane) i Wynik,min(zbadane), Wynik ost.
# ============================================================
test_min_map = compute_test_min_per_beam()  # Series: beam_key -> min_test (kN)

beams["P,min(zbadane) [kN]"] = beams["__beam_key"].map(test_min_map)
# DEBUG: pokaż jak Pandas widzi nagłówki (widać ukryte spacje/NBSP)def _build_beams_view(df_src: pd.DataFrame, geom_label: str) -> pd.DataFrame:
#     if df_src is None or df_src.empty:
#         return pd.DataFrame()
#
#     # --- mapowanie kolumn (elastyczne) ---
#     col_id = "ID" if "ID" in df_src.columns else None
#     col_name = "Nazwa belki" if "Nazwa belki" in df_src.columns else ("Nazwa" if "Nazwa" in df_src.columns else None)
#
#     # u Ciebie w arkuszu bywa "Receptura beton" (bez "u")
#     col_mix = None
#     for cand in ["Receptura betonu", "Receptura beton", "Mieszanka"]:
#         if cand in df_src.columns:
#             col_mix = cand
#             break
#
#     # geometria z arkusza (jeśli jest)
#     col_geom = "Geometria" if "Geometria" in df_src.columns else None
#
#     # cena mieszanki [USD/m³]
#     col_mix_price_m3 = None
#     for cand in [
#         "Cena mieszanki [USD/m³]",
#         "Cena mieszanki [USD/m3]",
#         "Cena mieszanki [USD/m^3]",
#         "Cena mieszanki / m3 [USD]",
#         "Cena mieszanki USD/m3",
#     ]:
#         if cand in df_src.columns:
#             col_mix_price_m3 = cand
#             break
#
#     # procedury
#     col_p_aci = "P_ACI_440_kN" if "P_ACI_440_kN" in df_src.columns else None
#     col_p_jsce = "P_JSCE_kN" if "P_JSCE_kN" in df_src.columns else None
#     col_p_csa = "P_CSA_kN" if "P_CSA_kN" in df_src.columns else None
#
#     # Pmin/własne + wyniki
#     col_p_min = "P_min_proc_kN" if "P_min_proc_kN" in df_src.columns else None
#     col_p_custom = "P_custom_kN" if "P_custom_kN" in df_src.columns else None
#     col_w_min = "Wynik_min_proc_USD_per_kN" if "Wynik_min_proc_USD_per_kN" in df_src.columns else None
#     col_w_custom = "Wynik_custom_USD_per_kN" if "Wynik_custom_USD_per_kN" in df_src.columns else None
#
#     # punktacja (kolumny “kosztowo-masowe”)
#     punkt_cols = [
#         "Łączna obj. belki [l]",
#         "Cena mieszanki / belkę [USD]",
#         "Łączna ilość prętów",
#         "Łączna cena zbrojenia [USD]",
#         "Całkowita masa belki [kg]",
#         "Koszt materiałów, brutto [USD]",
#         "Korekta materiałowa [%]",
#         "Koszt materiałów, netto [USD]",
#         "Korekta geometryczna [%]",
#         "Koszta transportu [USD]",
#         "Cena belki, brutto [USD]",
#         "Cena belki, netto [USD]",
#     ]
#
#     # jeśli w źródle ich nie ma, to je tworzę (żeby potem nie było KeyError)
#     for c in punkt_cols:
#         if c not in df_src.columns:
#             df_src[c] = ""
#
#     out = pd.DataFrame()
#     out["ID"] = df_src[col_id] if col_id else ""
#     out["Nazwa"] = df_src[col_name] if col_name else ""
#     out["Mieszanka"] = df_src[col_mix] if col_mix else ""
#     out["Cena mieszanki [USD/m³]"] = df_src[col_mix_price_m3] if col_mix_price_m3 else ""
#
#     # pokazuj tekst z arkusza (np. "prost."), ale kanonizuj do kluczy
#     if col_geom:
#         out["Geometria"] = df_src[col_geom].astype(str).fillna("").apply(lambda v: v.strip())
#     else:
#         out["Geometria"] = geom_label
#
#     out["__geom_id"] = out["Geometria"].apply(_canon_geom)
#
#     out["P_ACI_440 [kN]"] = df_src[col_p_aci] if col_p_aci else ""
#     out["P_JSCE [kN]"] = df_src[col_p_jsce] if col_p_jsce else ""
#     out["P_CSA [kN]"] = df_src[col_p_csa] if col_p_csa else ""
#
#     for c in punkt_cols:
#         out[c] = df_src[c]
#
#     out["P,min [kN]"] = df_src[col_p_min] if col_p_min else ""
#     out["P,własne [kN]"] = df_src[col_p_custom] if col_p_custom else ""
#     out["Wynik,min [USD/kN]"] = df_src[col_w_min] if col_w_min else ""
#     out["Wynik,własne [USD/kN]"] = df_src[col_w_custom] if col_w_custom else ""
#
#     # stabilny beam_key: używaj __geom_id (żeby "prost." -> "i")
#     def _mk_key(r):
#         rid = str(r.get("ID", "")).strip()
#         if rid:
#             return f"{r.get('__geom_id', '?')}|{rid}"
#         return f"{r.get('__geom_id', '?')}|{_normalize_name(r.get('Nazwa', ''))}"
#
#     out["__beam_key"] = out.apply(_mk_key, axis=1)
#     return out


# AWARYJNIE: jeżeli kolumny nie ma, utwórz ją pustą (żeby app nie padała)
if "Cena belki, netto [USD]" not in beams.columns:
    beams["Cena belki, netto [USD]"] = pd.NA
price_netto = beams["Cena belki, netto [USD]"].apply(to_num_pl)
pmin_zbad = beams["P,min(zbadane) [kN]"].apply(to_num_pl)

beams["Wynik,min(zbadane) [USD/kN]"] = (price_netto / pmin_zbad.where(pmin_zbad > 0)).replace([math.inf, -math.inf], pd.NA)
beams["Wynik ost. [USD/kN]"] = (beams["Wynik,min(zbadane) [USD/kN]"] * beams["Wsp. Raport"]).replace([math.inf, -math.inf], pd.NA)

# ============================================================
# Tabela główna (select + edycja Wsp. Raport)
# ============================================================
flt = st.text_input("Filtr (zawiera w nazwie belki):", "")
if flt.strip():
    beams = beams[beams["Nazwa"].astype(str).str.contains(flt.strip(), case=False, na=False)].copy()

if beams.empty:
    st.info("Brak belek po filtrze.")
    st.stop()

display_df = beams.copy()
display_df["__select__"] = False

cols_display = [
    "__select__",
    "ID",
    "Nazwa",
    "Mieszanka",
    "Cena mieszanki [USD/m³]",
    "Geometria",
    "P_ACI_440 [kN]",
    "P_JSCE [kN]",
    "P_CSA [kN]",
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
    "P,min [kN]",
    "P,własne [kN]",
    "Wynik,min [USD/kN]",
    "Wynik,własne [USD/kN]",
    "P,min(zbadane) [kN]",
    "Wynik,min(zbadane) [USD/kN]",
    "Wsp. Raport",
    "Wynik ost. [USD/kN]",
]

# disable wszystko poza __select__ i Wsp. Raport
disabled_cols = [c for c in cols_display if c not in ["__select__", "Wsp. Raport"]]

st.subheader("Tabela belek")
edited = st.data_editor(
    display_df[cols_display],
    use_container_width=True,
    hide_index=True,
    column_order=cols_display,
    column_config={
        "__select__": st.column_config.CheckboxColumn("", help="Zaznacz, aby zobaczyć wykonania/testy w expanderze."),
        "Cena mieszanki [USD/m³]": st.column_config.NumberColumn("Cena mieszanki [USD/m³]", format="%.2f"),
        "P_ACI_440 [kN]": st.column_config.NumberColumn("P_ACI_440 [kN]", format="%.2f"),
        "P_JSCE [kN]": st.column_config.NumberColumn("P_JSCE [kN]", format="%.2f"),
        "P_CSA [kN]": st.column_config.NumberColumn("P_CSA [kN]", format="%.2f"),
        "Łączna obj. belki [l]": st.column_config.NumberColumn("Łączna obj. belki [l]", format="%.1f"),
        "Cena mieszanki / belkę [USD]": st.column_config.NumberColumn("Cena mieszanki / belkę [USD]", format="%.2f"),
        "Łączna ilość prętów": st.column_config.NumberColumn("Łączna ilość prętów", format="%.0f"),
        "Łączna cena zbrojenia [USD]": st.column_config.NumberColumn("Łączna cena zbrojenia [USD]", format="%.2f"),
        "Całkowita masa belki [kg]": st.column_config.NumberColumn("Całkowita masa belki [kg]", format="%.2f"),
        "Koszt materiałów, brutto [USD]": st.column_config.NumberColumn("Koszt materiałów, brutto [USD]", format="%.2f"),
        "Korekta materiałowa [%]": st.column_config.NumberColumn("Korekta materiałowa [%]", format="%.0f"),
        "Koszt materiałów, netto [USD]": st.column_config.NumberColumn("Koszt materiałów, netto [USD]", format="%.2f"),
        "Korekta geometryczna [%]": st.column_config.NumberColumn("Korekta geometryczna [%]", format="%.0f"),
        "Koszta transportu [USD]": st.column_config.NumberColumn("Koszta transportu [USD]", format="%.2f"),
        "Cena belki, brutto [USD]": st.column_config.NumberColumn("Cena belki, brutto [USD]", format="%.2f"),
        "Cena belki, netto [USD]": st.column_config.NumberColumn("Cena belki, netto [USD]", format="%.2f"),
        "P,min [kN]": st.column_config.NumberColumn("P,min [kN]", format="%.2f"),
        "P,własne [kN]": st.column_config.NumberColumn("P,własne [kN]", format="%.2f"),
        "Wynik,min [USD/kN]": st.column_config.NumberColumn("Wynik,min [USD/kN]", format="%.2f"),
        "Wynik,własne [USD/kN]": st.column_config.NumberColumn("Wynik,własne [USD/kN]", format="%.2f"),
        "P,min(zbadane) [kN]": st.column_config.NumberColumn("P,min(zbadane) [kN]", format="%.2f"),
        "Wynik,min(zbadane) [USD/kN]": st.column_config.NumberColumn("Wynik,min(zbadane) [USD/kN]", format="%.2f"),
        "Wsp. Raport": st.column_config.NumberColumn("Wsp. Raport", help="Wprowadź współczynnik raportu", format="%.3f"),
        "Wynik ost. [USD/kN]": st.column_config.NumberColumn("Wynik ost. [USD/kN]", format="%.2f"),
    },
    disabled=disabled_cols,
)

# --- zapisz edytowane współczynniki do session_state map ---
try:
    edited_coef = edited[["Wsp. Raport"]].copy()
    for idx, val in edited_coef["Wsp. Raport"].items():
        beam_key = beams.loc[idx, "__beam_key"]
        v = to_num_pl(val)
        if pd.isna(v) or v is None or float(v) <= 0:
            v = 1.0
        coef_map[str(beam_key)] = float(v)
    st.session_state["report_coef_map"] = coef_map
except Exception:
    pass

selected_idx = edited.index[edited["__select__"] == True] if "__select__" in edited.columns else []



# ============================================================
# Expandery: wykonania + testy (bez wyboru typu)
# ============================================================
if len(selected_idx):
    st.subheader("Wykonania i testy")

    for irow in selected_idx:
        row = beams.loc[irow]
        beam_key = str(row["__beam_key"])
        name = str(row.get("Nazwa", ""))
        geom = str(row.get("Geometria", ""))
        bid = str(row.get("ID", ""))

        with st.expander(f"[{geom}] ID={bid} • {name}", expanded=False):
            # --- wykonania ---
            st.markdown("## Wykonania")

            load_exec_state_from_sheet(beam_key)

            exec_ns = f"exec__{beam_key}"
            EXEC_DF = f"{exec_ns}__df"
            EXEC_SNAP = f"{exec_ns}__snap"
            EXEC_HIST = f"{exec_ns}__hist"
            EXEC_WKEY = f"{exec_ns}__editor"

            exec_df = ensure_exec_ids(st.session_state[EXEC_DF]).copy()
            exec_df["Data wyk."] = exec_df["Data wyk."].astype("object").where(pd.notna(exec_df["Data wyk."]), "").astype(str)
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
                        st.toast("Usunięto wykonania.", icon="🗑️")
                        st.rerun()
                    else:
                        st.info("Nie zaznaczono żadnych wykonań do usunięcia (użyj kolumny 'Usuń?').")

            # auto-zapis zmian w session_state
            edited_no_ui = edited_exec.drop(columns=["_del"], errors="ignore").copy()
            for col_name in ["Data wyk.", "Wykonawca/y", "Uwagi"]:
                if col_name in edited_no_ui.columns:
                    edited_no_ui[col_name] = edited_no_ui[col_name].astype("object").where(pd.notna(edited_no_ui[col_name]), "")
            edited_no_ui = ensure_exec_ids(edited_no_ui)

            cur_snap = stable_json_exec(edited_no_ui)
            if cur_snap != st.session_state.get(EXEC_SNAP, ""):
                st.session_state[EXEC_HIST].append(st.session_state[EXEC_DF].copy())
                if len(st.session_state[EXEC_HIST]) > 30:
                    st.session_state[EXEC_HIST].pop(0)
                st.session_state[EXEC_DF] = edited_no_ui.copy()
                st.session_state[EXEC_SNAP] = cur_snap
                st.toast("Zapisano zmiany w wykonaniach (lokalnie).", icon="✅")
                st.rerun()

            current_exec_df = ensure_exec_ids(st.session_state[EXEC_DF]).reset_index(drop=True)

            # --- testy per wykonanie ---
            st.markdown("## Testy")

            for idx_exec, erow in current_exec_df.iterrows():
                nr_wyk = erow.get("Nr wyk.", idx_exec + 1)
                try:
                    nr_wyk_int = int(nr_wyk)
                except Exception:
                    nr_wyk_int = idx_exec + 1

                st.markdown(f"### Testy — Wykonanie {nr_wyk_int}")

                load_tests_state_from_sheet(beam_key, nr_wyk_int)

                tests_ns = f"tests__{beam_key}__{nr_wyk_int}"
                TDF = f"{tests_ns}__df"
                TSNAP = f"{tests_ns}__snap"
                TWKEY = f"{tests_ns}__editor"

                tests_df = ensure_test_ids(st.session_state[TDF].copy())
                tests_df["_del"] = False

                edited_tests = st.data_editor(
                    tests_df,
                    key=TWKEY,
                    num_rows="fixed",
                    use_container_width=True,
                    hide_index=True,
                    column_order=["_del", "Nr testu", "Data testu", "Wynik", "Wykonawca/y", "Uwagi"],
                    column_config={
                        "_del": st.column_config.CheckboxColumn("Usuń?", help="Zaznacz testy do usunięcia"),
                        "Nr testu": st.column_config.NumberColumn("Nr testu", disabled=True),
                        "Data testu": st.column_config.TextColumn("Data testu (DD-MM-YYYY)"),
                        "Wynik": st.column_config.TextColumn("Wynik [kN]"),
                        "Wykonawca/y": st.column_config.TextColumn("Wykonawca/y"),
                        "Uwagi": st.column_config.TextColumn("Uwagi"),
                    },
                )

                c_add_test, c_del_test, _ = st.columns([1, 1, 4])

                with c_add_test:
                    if st.button("➕ Dodaj test", key=f"{tests_ns}__add"):
                        cur_tests = st.session_state[TDF]
                        new_row_t = {"Nr testu": None, "Data testu": "", "Wynik": "", "Wykonawca/y": "", "Uwagi": ""}
                        new_tests_df = pd.concat([cur_tests, pd.DataFrame([new_row_t])], ignore_index=True)
                        new_tests_df = ensure_test_ids(new_tests_df)
                        st.session_state[TDF] = new_tests_df
                        st.session_state[TSNAP] = stable_json_test(new_tests_df)
                        st.rerun()

                with c_del_test:
                    if st.button("🗑️ Usuń zaznaczone testy", key=f"{tests_ns}__del"):
                        if "_del" in edited_tests and bool(edited_tests["_del"].any()):
                            mask_keep_t = ~edited_tests["_del"].astype(bool)
                            new_tests = edited_tests.loc[mask_keep_t].drop(columns=["_del"])
                            new_tests = ensure_test_ids(new_tests)

                            st.session_state[TDF] = new_tests.copy()
                            st.session_state[TSNAP] = stable_json_test(new_tests)
                            st.toast("Usunięto testy.", icon="🗑️")
                            st.rerun()
                        else:
                            st.info("Nie zaznaczono żadnych testów do usunięcia (użyj kolumny 'Usuń?').")

                # auto-zapis zmian testów w session_state
                edited_tests_no_ui = edited_tests.drop(columns=["_del"], errors="ignore").copy()
                for col_name in ["Data testu", "Wynik", "Wykonawca/y", "Uwagi"]:
                    if col_name in edited_tests_no_ui.columns:
                        edited_tests_no_ui[col_name] = edited_tests_no_ui[col_name].astype("object").where(pd.notna(edited_tests_no_ui[col_name]), "")
                edited_tests_no_ui = ensure_test_ids(edited_tests_no_ui)

                cur_t_snap = stable_json_test(edited_tests_no_ui)
                if cur_t_snap != st.session_state.get(TSNAP, ""):
                    st.session_state[TDF] = edited_tests_no_ui.copy()
                    st.session_state[TSNAP] = cur_t_snap
                    st.toast("Zapisano zmiany w testach (lokalnie).", icon="✅")
                    st.rerun()

