import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(layout="wide")

HISTORY_DIR = Path("history")
HISTORY_DIR.mkdir(exist_ok=True)

# -------------------------
# HELPER
# -------------------------
def normalize_code(x):
    return str(x).strip()

def excel_serial_to_date(val):
    try:
        return pd.to_datetime(val)
    except:
        return pd.NaT

def get_week_start(date):
    return pd.to_datetime(date).to_period("W").start_time

# -------------------------
# SUPPLY OKUMA
# -------------------------
def read_supply_file(file):

    xls = pd.ExcelFile(file)
    sheet = xls.sheet_names[0]

    raw = pd.read_excel(file, sheet_name=sheet, header=None)

    header_row = None
    for i in range(min(10, len(raw))):
        row = raw.iloc[i]
        if sum(pd.to_datetime(v, errors="coerce") is not pd.NaT for v in row) >= 2:
            header_row = i
            break

    if header_row is None:
        header_row = 0

    headers = list(raw.iloc[header_row])
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers

    df = df.dropna(how="all")

    code_col = df.columns[0]

    df = df.rename(columns={code_col: "product_code"})
    df["product_code"] = df["product_code"].apply(normalize_code)

    date_cols = [c for c in df.columns if c != "product_code"]

    df = df.melt(
        id_vars=["product_code"],
        value_vars=date_cols,
        var_name="date",
        value_name="inbound_qty"
    )

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["inbound_qty"] = pd.to_numeric(df["inbound_qty"], errors="coerce").fillna(0)

    df = df.dropna(subset=["date"])
    df = df[df["inbound_qty"] != 0]

    df["date"] = df["date"] - pd.Timedelta(days=7)
    df["week"] = df["date"].dt.to_period("W").astype(str)

    return df

# -------------------------
# APO OKUMA
# -------------------------
def read_apo_file(file):

    xls = pd.ExcelFile(file)
    sheet = xls.sheet_names[0]

    raw = pd.read_excel(file, sheet_name=sheet, header=None)

    header_row = 1
    headers = list(raw.iloc[header_row])

    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers

    code_col = df.columns[0]

    df = df.rename(columns={code_col: "product_code"})
    df["product_code"] = df["product_code"].apply(normalize_code)

    date_cols = df.columns[1:]

    df = df.melt(
        id_vars=["product_code"],
        value_vars=date_cols,
        var_name="date",
        value_name="outbound_qty"
    )

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["outbound_qty"] = pd.to_numeric(df["outbound_qty"], errors="coerce").fillna(0)

    df = df.dropna(subset=["date"])
    df = df[df["outbound_qty"] != 0]

    df["week"] = df["date"].dt.to_period("W").astype(str)

    return df

# -------------------------
# EXCEL FORMAT
# -------------------------
def format_excel(data):

    bio = BytesIO()
    data.to_excel(bio, index=False)

    bio.seek(0)
    wb = load_workbook(bio)

    ws = wb.active

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# -------------------------
# UI
# -------------------------
st.title("📦 Depo Dashboard")

supply_file = st.file_uploader("Supply dosyası")
apo_file = st.file_uploader("APO dosyası")

col1, col2, col3 = st.columns(3)

with col1:
    depo_kapasite = st.number_input("Depo kapasite", value=1100)

with col2:
    takip = st.number_input("Takip eşiği", value=85)

with col3:
    kritik = st.number_input("Kritik eşiği", value=99)

if st.button("Hesapla"):

    supply = read_supply_file(supply_file)
    apo = read_apo_file(apo_file)

    df = pd.merge(
        supply,
        apo,
        on=["product_code", "week"],
        how="outer"
    ).fillna(0)

    weekly = df.groupby("week")[["inbound_qty", "outbound_qty"]].sum().reset_index()

    weekly["stock"] = weekly["inbound_qty"].cumsum() - weekly["outbound_qty"].cumsum()

    weekly["palet"] = (weekly["stock"] / 2400).round(0)

    weekly["kapasite_%"] = (weekly["palet"] / depo_kapasite) * 100

    st.dataframe(
        weekly.style.format({
            "inbound_qty": "{:,.0f}",
            "outbound_qty": "{:,.0f}",
            "stock": "{:,.0f}",
            "palet": "{:,.0f}",
            "kapasite_%": "{:,.0f}"
        })
    )

    excel_bytes = format_excel(weekly)

    st.session_state["report"] = excel_bytes

    st.download_button(
        "Excel indir",
        data=excel_bytes,
        file_name="rapor.xlsx"
    )

# -------------------------
# SAVE
# -------------------------
st.subheader("Kaydet")

if "report" in st.session_state:

    name = st.text_input("Rapor adı", "depo_rapor")

    if st.button("Kaydet"):
        path = HISTORY_DIR / f"{name}.xlsx"
        path.write_bytes(st.session_state["report"])
        st.success("Kaydedildi")

# -------------------------
# HISTORY
# -------------------------
st.subheader("Geçmiş")

files = list(HISTORY_DIR.glob("*.xlsx"))

if files:

    selected = st.selectbox("Seç", files)

    if st.button("Aç"):
        df = pd.read_excel(selected)
        st.dataframe(df)
