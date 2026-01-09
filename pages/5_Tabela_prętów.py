import streamlit as st
import pandas as pd
import time

# === KONFIG GOOGLE SHEETS (opcjonalna; przyciski aktywne jeśli skonfigurowano secrets) ===
GS_READY = False
try:
    from google.oauth2.service_account import Credentials
    import gspread
    try:
        from gspread_dataframe import set_with_dataframe, get_as_dataframe
    except Exception:
        set_with_dataframe = None
        get_as_dataframe = None

    GSA = "gcp_service_account"
    SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")
    SHEET_GFRP = st.secrets.get("SHEET_GFRP", "gfrp_bars")  # osobny arkusz dla GFRP

    if GSA in st.secrets and SPREADSHEET_ID:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        CREDS = Credentials.from_service_account_info(st.secrets[GSA], scopes=SCOPES)
        GS_READY = True
except Exception:
    GS_READY = False

st.set_page_config(page_title="Pręty GFRP – baza", layout="wide")

# --- unikalna przestrzeń nazw dla TEJ strony/edytora ---
NS = "gfrp"
DATA   = f"{NS}__df"
SNAP   = f"{NS}__snapshot"
HIST   = f"{NS}__history"
WKEY   = f"{NS}__editor"

# --- stałe dla wyboru profilu i jednostek ceny ---
PROFILES = ["gładki", "spiralny", "piaskowany", "żebrowany"]
PRICE_UNITS = ["mb", "kg"]  # najczęściej metr bieżący

def stable_json(df: pd.DataFrame) -> str:
    if df is None or len(df) == 0:
        return "[]"
    return df.sort_index().sort_index(axis=1).to_json(
        orient="records", date_format="iso", date_unit="s"
    )

def ensure_ids(df: pd.DataFrame, id_col: str = "id") -> pd.DataFrame:
    if df is None or id_col not in df.columns:
        return df
    out = df.copy()
    s = pd.to_numeric(out[id_col], errors="coerce")
    miss = s.isna()
    current_max = s.dropna().max()
    if pd.isna(current_max):
        current_max = 0
    if miss.any():
        n_missing = int(miss.sum())
        start = int(current_max) + 1
        out.loc[miss, id_col] = list(range(start, start + n_missing))
    out[id_col] = pd.to_numeric(out[id_col], errors="coerce").astype("Int64")
    return out

# ==== Funkcje Google Sheets ====
def wanted_cols_for_sheet(df: pd.DataFrame):
    # kolumny zapisywane do arkusza
    cols = [
        "id",
        "nazwa",
        "srednica_mm",
        "profil",
        "R_t_MPa",
        "E_GPa",
        "τ_base_MPa",
        "gestosc_gcm3",
        "cena_pln",
        "cena_za",
        "co2e_kgkg",
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = None
    return cols

def _open_ws():
    assert GS_READY, "Brak konfiguracji Google Sheets"
    gc = gspread.authorize(CREDS)
    ss = gc.open_by_key(SPREADSHEET_ID)
    try:
        return ss.worksheet(SHEET_GFRP)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=SHEET_GFRP, rows=1000, cols=20)
        ws.update("A1", [[
            "id", "nazwa", "srednica_mm", "profil",
            "R_t_MPa", "E_GPa", "τ_base_MPa", "gestosc_gcm3",
            "cena_pln", "cena_za", "co2e_kgkg"
        ]])
        return ws

def read_gfrp_from_sheet() -> pd.DataFrame:
    ws = _open_ws()
    if 'get_as_dataframe' in globals() and get_as_dataframe is not None:
        df = get_as_dataframe(ws, evaluate_formulas=True, header=0).dropna(how="all")
    else:
        rows = ws.get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(rows)

    wanted = [
        "id", "nazwa", "srednica_mm", "profil",
        "R_t_MPa", "E_GPa","τ_base_MPa", "gestosc_gcm3",
        "cena_pln", "cena_za", "co2e_kgkg"
    ]
    for c in wanted:
        if c not in df.columns:
            df[c] = None

    # liczby
    for c in ["id", "srednica_mm", "R_t_MPa", "E_GPa", "τ_base_MPa", "gestosc_gcm3", "cena_pln", "co2e_kgkg"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["profil"] = df["profil"].astype(object)
    df["cena_za"] = df["cena_za"].astype(object)

    # <<< BEZPIECZNIK: jeśli arkusz jest pusty -> jeden wiersz startowy
    if df is None or len(df.dropna(how="all")) == 0:
        df = _default_df()

    return df[wanted]

def write_gfrp_to_sheet(df: pd.DataFrame):
    ws = _open_ws()
    cols = wanted_cols_for_sheet(df)
    if 'set_with_dataframe' in globals() and set_with_dataframe is not None:
        ws.clear()
        ws_df = df[cols]
        set_with_dataframe(ws, ws_df, include_index=False, include_column_header=True, resize=True)
    else:
        ws.clear()
        values = [cols] + df[cols].astype(object).where(pd.notnull(df[cols]), "").values.tolist()
        ws.update("A1", values)

# ---- HYDRACJA: przy pierwszym uruchomieniu wczytaj z Google Sheets (jeśli możliwe) ----
def _default_df() -> pd.DataFrame:
    # <<< START: dokładnie jeden wiersz, id=1
    return pd.DataFrame({
        "id": [1],
        "nazwa": [""],
        "srednica_mm": [None],
        "profil": [None],
        "R_t_MPa": [None],
        "E_GPa": [None],
        "τ_base_MPa": [None],
        "gestosc_gcm3": [None],
        "cena_pln": [None],
        "cena_za": [None],
        "co2e_kgkg": [None],
    })

if DATA not in st.session_state:
    if GS_READY:
        try:
            df_gs = ensure_ids(read_gfrp_from_sheet(), "id")
            # <<< jeśli z jakiegoś powodu nadal pusto (np. same NaN) -> wiersz startowy
            if df_gs is None or len(df_gs.dropna(how="all")) == 0:
                df_gs = _default_df()
            st.session_state[DATA] = df_gs.copy()
            st.session_state[SNAP] = stable_json(df_gs)
            st.session_state[HIST] = []
            st.toast("Wczytano pręty GFRP z Google Sheets ✔️", icon="✅")
        except Exception as e:
            st.session_state[DATA] = ensure_ids(_default_df(), "id")
            st.session_state[SNAP] = stable_json(st.session_state[DATA])
            st.session_state[HIST] = []
            st.warning(f"Nie udało się wczytać GFRP z Google Sheets na starcie: {e}. Używam wiersza startowego.")
    else:
        st.session_state[DATA] = ensure_ids(_default_df(), "id")
        st.session_state[SNAP] = stable_json(st.session_state[DATA])
        st.session_state[HIST] = []
else:
    df = st.session_state[DATA]
    # Migracje (gdybyś coś zmieniał potem)
    if "cena_za" not in df.columns:
        df = df.assign(cena_za=None)
    if "co2e_kgkg" not in df.columns:
        df = df.assign(co2e_kgkg=None)
    st.session_state[DATA] = df

st.title("Baza prętów GFRP")

# --- wspólny layout
LAYOUT = [1, 1, 8]

# === GÓRNY pasek akcji ===
c_save, c_update, _sp_top = st.columns(LAYOUT)
with c_save:
    if st.button("💾 Zapisz", use_container_width=True):
        if not GS_READY:
            st.error("Brak konfiguracji Google Sheets w secrets. Uzupełnij SPREADSHEET_ID i [gcp_service_account].")
        else:
            try:
                df_to_save = ensure_ids(st.session_state[DATA].copy(), "id")
                write_gfrp_to_sheet(df_to_save)
                st.session_state[SNAP] = stable_json(st.session_state[DATA])
                st.toast("Zapisano pręty GFRP do Google Sheets ✔️", icon="✅")
                time.sleep(0.3)
                st.rerun()
            except Exception as e:
                st.error(f"Nie udało się zapisać do Google Sheets: {e}")

with c_update:
    if st.button("↻ Zaktualizuj", use_container_width=True):
        if not GS_READY:
            st.error("Brak konfiguracji Google Sheets w secrets. Uzupełnij SPREADSHEET_ID i [gcp_service_account].")
        else:
            try:
                df_gs = ensure_ids(read_gfrp_from_sheet(), "id")
                # <<< BEZPIECZNIK po aktualizacji: jak pusto -> wiersz startowy
                if df_gs is None or len(df_gs.dropna(how="all")) == 0:
                    df_gs = _default_df()
                st.session_state[HIST].append(st.session_state[DATA].copy())
                if len(st.session_state[HIST]) > 30:
                    st.session_state[HIST].pop(0)
                st.session_state[DATA] = df_gs.copy()
                st.session_state[SNAP] = stable_json(df_gs)
                st.toast("Wczytano GFRP z Google Sheets 🔄", icon="🔄")
                st.rerun()
            except Exception as e:
                st.error(f"Nie udało się wczytać z Google Sheets: {e}")

before = stable_json(st.session_state[DATA])

# --- edytor: KOPIA + kolumny pomocnicze
df_for_edit = st.session_state[DATA].copy()
df_for_edit.insert(0, "_rowid", df_for_edit.index.astype(int))
df_for_edit.insert(1, "_del", False)

# kosmetyka: puste pola zamiast 'None' dla tekstów
for col in ["nazwa", "profil", "cena_za"]:
    if col in df_for_edit.columns:
        df_for_edit[col] = df_for_edit[col].astype("object").where(pd.notna(df_for_edit[col]), "")

edited = st.data_editor(
    df_for_edit,
    key=WKEY,
    num_rows="fixed",   # <-- wyłączamy systemowe dodawanie
    use_container_width=True,
    hide_index=True,
    column_order=[
        "_del", "id", "nazwa",
        "srednica_mm", "profil",
        "R_t_MPa", "E_GPa", "τ_base_MPa",
        "gestosc_gcm3",
        "cena_pln", "cena_za",
        "co2e_kgkg",
    ],
    column_config={
        "_del": st.column_config.CheckboxColumn("Usuń?", help="Zaznacz wiersze do usunięcia"),
        "id": st.column_config.NumberColumn("ID", disabled=True),
        "nazwa": st.column_config.TextColumn("Nazwa", max_chars=200),
        "srednica_mm": st.column_config.NumberColumn("⌀ [mm]", step=1),
        "profil": st.column_config.SelectboxColumn("Profil", options=PROFILES, default="spiralny"),
        "R_t_MPa": st.column_config.NumberColumn("Rₜ [MPa]", step=10),
        "E_GPa": st.column_config.NumberColumn("E [GPa]", step=1),
        "τ_base_MPa": st.column_config.NumberColumn("τ_base [MPa]", step=1),
        "gestosc_gcm3": st.column_config.NumberColumn("Gęstość [g/cm³]", step=0.01, format="%.2f"),
        "cena_pln": st.column_config.NumberColumn("Cena [PLN]", format="%.2f"),
        "cena_za": st.column_config.SelectboxColumn("Cena za", options=PRICE_UNITS, default="mb"),
        "co2e_kgkg": st.column_config.NumberColumn("CO₂e [kg/kg]", step=0.001, format="%.3f"),
    },
)


# === DOLNY pasek akcji ===
c_delete, c_add, _sp_bottom = st.columns(LAYOUT)

with c_delete:
    if st.button("🗑️ Usuń zaznaczone", use_container_width=True):
        if "_del" in edited and bool(edited["_del"].any()):
            rowids_to_delete = edited.loc[edited["_del"] == True, "_rowid"].dropna().astype(int).tolist()
            if rowids_to_delete:
                new_df = st.session_state[DATA].drop(index=rowids_to_delete, errors="ignore").copy()
                st.session_state[DATA] = new_df
                st.session_state[SNAP] = stable_json(new_df)
                st.toast(f"Usunięto {len(rowids_to_delete)} wiersz(e).", icon="🗑️")
                st.rerun()
        else:
            st.warning("Nie zaznaczono żadnego wiersza do usunięcia.")

with c_add:
    if st.button("➕ Dodaj nowy wiersz", use_container_width=True):
        df = st.session_state[DATA].copy()
        new_id = int(df["id"].max()) + 1 if len(df) else 1
        new_row = {
            "id": new_id,
            "nazwa": "",
            "srednica_mm": None,
            "profil": None,
            "R_t_MPa": None,
            "E_GPa": None,
            "τ_base_MPa": None,
            "gestosc_gcm3": None,
            "cena_pln": None,
            "cena_za": None,
            "co2e_kgkg": None,
        }
        st.session_state[DATA] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        st.session_state[SNAP] = stable_json(st.session_state[DATA])
        st.toast("Dodano nowy wiersz ✔️", icon="➕")
        st.rerun()


# --- przygotuj wersję bez kolumn pomocniczych do walidacji/autosave ---
edited_no_ui = edited.drop(columns=["_del", "_rowid"], errors="ignore").copy()

# typowanie pól liczbowych
for col in ["srednica_mm", "R_t_MPa", "E_GPa", "τ_base_MPa", "gestosc_gcm3", "cena_pln", "co2e_kgkg"]:
    if col in edited_no_ui:
        edited_no_ui[col] = pd.to_numeric(edited_no_ui[col], errors="coerce")

if "profil" in edited_no_ui:
    edited_no_ui["profil"] = edited_no_ui["profil"].where(
        edited_no_ui["profil"].isin(PROFILES), None
    )

if "cena_za" in edited_no_ui:
    edited_no_ui["cena_za"] = edited_no_ui["cena_za"].where(
        edited_no_ui["cena_za"].isin(PRICE_UNITS), None
    )

edited_no_ui = ensure_ids(edited_no_ui, "id")

# --- AUTOSAVE ---
current = stable_json(edited_no_ui)
if current != st.session_state[SNAP]:
    st.session_state[HIST].append(st.session_state[DATA].copy())
    if len(st.session_state[HIST]) > 30:
        st.session_state[HIST].pop(0)
    st.session_state[DATA] = edited_no_ui.copy()
    st.session_state[SNAP] = current
    st.toast("Autozapisano ✔️ (lokalnie). Użyj 💾 aby wysłać do Google Sheets.", icon="✅")
    st.rerun()

# --- diagnoza nadpisań ---
after = stable_json(st.session_state[DATA])
if before != st.session_state[SNAP] and before != after:
    st.error("🔴 Wykryto nadpisanie stanu PRZED edytorem (jakiś kod zresetował dane).")


