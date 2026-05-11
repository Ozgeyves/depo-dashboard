import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="Depo Dashboard", layout="wide")

HISTORY_DIR = Path("history")
HISTORY_DIR.mkdir(exist_ok=True)

# =========================================================
# LOGIN
# =========================================================
PASSWORDS = {
    "tedarik": "tedarik",
    "depo": "depo",
}


def login_screen():
    st.title("📦 Depo Dashboard")
    st.caption("Giriş yapmak için şifrenizi yazınız.")

    password = st.text_input("Şifre", type="password")
    login_button = st.button("Giriş Yap")

    if login_button:
        role = PASSWORDS.get(str(password).strip().lower())
        if role:
            st.session_state["logged_in"] = True
            st.session_state["role"] = role
            st.rerun()
        else:
            st.error("Şifre hatalı. Lütfen tekrar deneyin.")


def logout_button():
    with st.sidebar:
        if st.button("Çıkış Yap"):
            st.session_state.clear()
            st.rerun()


if not st.session_state.get("logged_in"):
    login_screen()
    st.stop()

ROLE = st.session_state.get("role", "depo")
IS_TEDARIK = ROLE == "tedarik"

# =========================================================
# HELPERS
# =========================================================

def normalize_code(x):
    if pd.isna(x):
        return ""
    return str(x).strip().replace(".0", "")


def clean_number(value):
    if pd.isna(value):
        return 0
    if isinstance(value, str):
        value = value.replace(".", "").replace(",", ".")
    return pd.to_numeric(value, errors="coerce") if not pd.isna(value) else 0


def safe_to_datetime(value):
    return pd.to_datetime(value, errors="coerce", dayfirst=True)


def week_label(date_series):
    return pd.to_datetime(date_series, errors="coerce").dt.to_period("W-MON").astype(str)


def find_header_row_with_dates(raw: pd.DataFrame, max_rows: int = 15) -> int:
    for i in range(min(max_rows, len(raw))):
        parsed_dates = pd.to_datetime(raw.iloc[i], errors="coerce", dayfirst=True)
        if parsed_dates.notna().sum() >= 2:
            return i
    return 0


def first_non_empty_column(df: pd.DataFrame):
    for col in df.columns:
        if df[col].notna().sum() > 0:
            return col
    return df.columns[0]


def to_excel_download(df_dict: dict) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            safe_sheet = str(sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe_sheet)
    output.seek(0)
    return output


def save_history(weekly_df: pd.DataFrame, detail_df: pd.DataFrame):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = HISTORY_DIR / f"depo_dashboard_{timestamp}.xlsx"
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        weekly_df.to_excel(writer, index=False, sheet_name="Haftalik Ozet")
        detail_df.to_excel(writer, index=False, sheet_name="Urun Detay")
    return file_path

# =========================================================
# FILE READERS
# =========================================================

def read_supply_file(file):
    if file is None:
        return pd.DataFrame(columns=["product_code", "date", "inbound_qty", "week"])

    xls = pd.ExcelFile(file)
    sheet = xls.sheet_names[0]
    raw = pd.read_excel(file, sheet_name=sheet, header=None)

    header_row = find_header_row_with_dates(raw)
    headers = list(raw.iloc[header_row])
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all")

    code_col = first_non_empty_column(df)
    df = df.rename(columns={code_col: "product_code"})
    df["product_code"] = df["product_code"].apply(normalize_code)
    df = df[df["product_code"] != ""]

    date_cols = [c for c in df.columns if c != "product_code"]

    melted = df.melt(
        id_vars=["product_code"],
        value_vars=date_cols,
        var_name="date",
        value_name="inbound_qty",
    )

    melted["date"] = pd.to_datetime(melted["date"], errors="coerce", dayfirst=True)
    melted["inbound_qty"] = pd.to_numeric(melted["inbound_qty"], errors="coerce").fillna(0)
    melted = melted.dropna(subset=["date"])
    melted = melted[melted["inbound_qty"] != 0]

    # Önceki kod mantığı korunuyor: Supply tarihi 7 gün geri çekiliyor.
    melted["date"] = melted["date"] - pd.Timedelta(days=7)
    melted["week"] = week_label(melted["date"])

    return melted


def read_apo_file(file):
    if file is None:
        return pd.DataFrame(columns=["product_code", "date", "outbound_qty", "week"])

    xls = pd.ExcelFile(file)
    sheet = xls.sheet_names[0]
    raw = pd.read_excel(file, sheet_name=sheet, header=None)

    # Önceki kod mantığı korunuyor: APO başlık satırı 2. satır kabul ediliyor.
    header_row = 1 if len(raw) > 1 else 0
    headers = list(raw.iloc[header_row])
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all")

    code_col = first_non_empty_column(df)
    df = df.rename(columns={code_col: "product_code"})
    df["product_code"] = df["product_code"].apply(normalize_code)
    df = df[df["product_code"] != ""]

    date_cols = [c for c in df.columns if c != "product_code"]

    melted = df.melt(
        id_vars=["product_code"],
        value_vars=date_cols,
        var_name="date",
        value_name="outbound_qty",
    )

    melted["date"] = pd.to_datetime(melted["date"], errors="coerce", dayfirst=True)
    melted["outbound_qty"] = pd.to_numeric(melted["outbound_qty"], errors="coerce").fillna(0)
    melted = melted.dropna(subset=["date"])
    melted = melted[melted["outbound_qty"] != 0]
    melted["week"] = week_label(melted["date"])

    return melted


def read_depo_data_file(file):
    """Depo data dosyası standart değilse bile okunabilir bilgi tablosu üretir."""
    if file is None:
        return pd.DataFrame()

    xls = pd.ExcelFile(file)
    frames = []

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(file, sheet_name=sheet)
            df = df.dropna(how="all")
            if df.empty:
                continue
            df["sheet_name"] = sheet
            frames.append(df)
        except Exception:
            continue

    if not frames:
        return pd.DataFrame()

    depo = pd.concat(frames, ignore_index=True)

    # Kolon isimlerini daha okunur hale getirir.
    depo.columns = [str(c).strip() for c in depo.columns]

    # Ürün kodu benzeri kolon varsa normalize eder.
    possible_code_cols = [c for c in depo.columns if any(k in c.lower() for k in ["kod", "code", "product", "ürün", "urun"])]
    if possible_code_cols:
        depo[possible_code_cols[0]] = depo[possible_code_cols[0]].apply(normalize_code)

    return depo

# =========================================================
# UI
# =========================================================

st.title("📦 Depo Giriş - Çıkış / Stok Seviye Raporu")
logout_button()

with st.sidebar:
    st.header("Parametreler")

    baslangic_ana = st.number_input("Başlangıç Ana Ürün", value=0, step=1)
    baslangic_mini = st.number_input("Başlangıç Mini Sample", value=0, step=1)
    sarf = st.number_input("Haftalık Sarf Palet", value=250, step=1)

    if IS_TEDARIK:
        st.divider()
        st.subheader("Hesap Parametreleri")
        ana_palet = st.number_input("Ana Ürün Palet İçi", value=2400, step=1)
        mini_palet = st.number_input("Mini Sample Palet İçi", value=15000, step=1)
        adr_palet = st.number_input("ADR Palet İçi", value=5540, step=1)
        tir_kapasite = st.number_input("1 Tır Kaç Palet", value=40, step=1)
    else:
        ana_palet = 2400
        mini_palet = 15000
        adr_palet = 5540
        tir_kapasite = 40

st.subheader("Dosya Yükleme")
col1, col2, col3 = st.columns(3)
with col1:
    supply_file = st.file_uploader("Supply dosyası", type=["xlsx", "xls", "xlsm"], key="supply")
with col2:
    apo_file = st.file_uploader("APO dosyası", type=["xlsx", "xls", "xlsm"], key="apo")
with col3:
    depo_file = st.file_uploader("Depo data dosyası", type=["xlsx", "xls", "xlsm"], key="depo_data")

hesapla = st.button("Hesapla", type="primary")

if hesapla:
    try:
        if supply_file is None and apo_file is None and depo_file is None:
            st.warning("Lütfen en az bir dosya yükleyin.")
            st.stop()

        supply = read_supply_file(supply_file)
        apo = read_apo_file(apo_file)
        depo_data = read_depo_data_file(depo_file)

        detail = pd.merge(
            supply[["product_code", "week", "date", "inbound_qty"]],
            apo[["product_code", "week", "date", "outbound_qty"]],
            on=["product_code", "week"],
            how="outer",
            suffixes=("_supply", "_apo"),
        )

        if detail.empty:
            detail = pd.DataFrame(columns=["product_code", "week", "inbound_qty", "outbound_qty"])

        detail["inbound_qty"] = pd.to_numeric(detail.get("inbound_qty", 0), errors="coerce").fillna(0)
        detail["outbound_qty"] = pd.to_numeric(detail.get("outbound_qty", 0), errors="coerce").fillna(0)
        detail["net_qty"] = detail["inbound_qty"] - detail["outbound_qty"]

        weekly = detail.groupby("week", as_index=False)[["inbound_qty", "outbound_qty", "net_qty"]].sum()
        weekly = weekly.sort_values("week")

        starting_stock = baslangic_ana + baslangic_mini
        weekly["opening_stock"] = starting_stock + weekly["net_qty"].cumsum().shift(1).fillna(0)
        weekly["closing_stock"] = starting_stock + weekly["net_qty"].cumsum()

        if IS_TEDARIK:
            weekly["palet"] = (weekly["closing_stock"] / ana_palet).replace([np.inf, -np.inf], 0).fillna(0)
            weekly["tir"] = (weekly["palet"] / tir_kapasite).replace([np.inf, -np.inf], 0).fillna(0)
            weekly["sarf_sonrasi_palet"] = weekly["palet"] - sarf

        st.success("Rapor oluşturuldu.")

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        total_in = weekly["inbound_qty"].sum() if not weekly.empty else 0
        total_out = weekly["outbound_qty"].sum() if not weekly.empty else 0
        last_stock = weekly["closing_stock"].iloc[-1] if not weekly.empty else starting_stock
        week_count = weekly["week"].nunique() if not weekly.empty else 0

        kpi1.metric("Toplam Giriş", f"{total_in:,.0f}")
        kpi2.metric("Toplam Çıkış", f"{total_out:,.0f}")
        kpi3.metric("Son Kapanış Stok", f"{last_stock:,.0f}")
        kpi4.metric("Hafta Sayısı", f"{week_count:,.0f}")

        st.subheader("Haftalık Özet")

        if IS_TEDARIK:
            show_cols = [
                "week", "opening_stock", "inbound_qty", "outbound_qty", "net_qty",
                "closing_stock", "palet", "tir", "sarf_sonrasi_palet"
            ]
            format_dict = {
                "opening_stock": "{:,.0f}",
                "inbound_qty": "{:,.0f}",
                "outbound_qty": "{:,.0f}",
                "net_qty": "{:,.0f}",
                "closing_stock": "{:,.0f}",
                "palet": "{:,.0f}",
                "tir": "{:,.1f}",
                "sarf_sonrasi_palet": "{:,.0f}",
            }
        else:
            show_cols = ["week", "opening_stock", "inbound_qty", "outbound_qty", "net_qty", "closing_stock"]
            format_dict = {
                "opening_stock": "{:,.0f}",
                "inbound_qty": "{:,.0f}",
                "outbound_qty": "{:,.0f}",
                "net_qty": "{:,.0f}",
                "closing_stock": "{:,.0f}",
            }

        st.dataframe(
            weekly[show_cols].style.format(format_dict),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Ürün Bazlı Detay")
        detail_show = detail.copy()
        detail_show = detail_show.sort_values(["week", "product_code"])
        detail_cols = [c for c in ["product_code", "week", "inbound_qty", "outbound_qty", "net_qty"] if c in detail_show.columns]
        st.dataframe(
            detail_show[detail_cols].style.format({
                "inbound_qty": "{:,.0f}",
                "outbound_qty": "{:,.0f}",
                "net_qty": "{:,.0f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        if depo_file is not None:
            st.subheader("Depo Data Bilgi Alanı")
            if depo_data.empty:
                st.info("Depo data dosyası okundu ancak gösterilecek dolu veri bulunamadı.")
            else:
                st.dataframe(depo_data, use_container_width=True, hide_index=True)

        sheets = {
            "Haftalik Ozet": weekly[show_cols],
            "Urun Detay": detail_show[detail_cols],
        }
        if depo_file is not None and not depo_data.empty:
            sheets["Depo Data"] = depo_data

        excel_file = to_excel_download(sheets)
        st.download_button(
            label="Excel Olarak İndir",
            data=excel_file,
            file_name=f"depo_dashboard_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        try:
            save_history(weekly[show_cols], detail_show[detail_cols])
        except Exception:
            pass

    except Exception as e:
        st.error("Rapor oluşturulurken hata oluştu.")
        st.exception(e)
else:
    st.info("Dosyaları yükleyip Hesapla butonuna basın.")
