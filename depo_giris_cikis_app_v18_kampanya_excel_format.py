import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pathlib import Path

# -----------------------------
# CONFIG
# -----------------------------
st.set_page_config(layout="wide", page_title="Depo Dashboard")

HISTORY_DIR = Path("history")
HISTORY_DIR.mkdir(exist_ok=True)

# -----------------------------
# HELPERS
# -----------------------------
def normalize_code(x):
    return str(x).strip()

def get_week_start(date):
    return pd.to_datetime(date).to_period("W").start_time

# -----------------------------
# SUPPLY OKUMA (FIXED)
# -----------------------------
def read_supply_file(file):

    xls = pd.ExcelFile(file)
    sheet = xls.sheet_names[0]

    raw = pd.read_excel(file, sheet_name=sheet, header=None)

    header_row = 0
    for i in range(min(10, len(raw))):
        row = raw.iloc[i]
        date_count = sum(pd.to_datetime(v, errors="coerce") is not pd.NaT for v in row)
        if date_count >= 2:
            header_row = i
            break

    headers = list(raw.iloc[header_row])
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all")

    # ürün kodu bul
    code_col = df.columns[0]

    df = df.rename(columns={code_col: "product_code"})
    df["product_code"] = df["product_code"].astype(str)

    date_cols = [c for c in df.columns if c != "product_code"]

    df = df.melt(
        id_vars=["product_code"],
        value_vars=date_cols,
        var_name="date",
        value_name="inbound"
    )

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["inbound"] = pd.to_numeric(df["inbound"], errors="coerce").fillna(0)

    df = df.dropna(subset=["date"])
    df = df[df["inbound"] != 0]

    df["date"] = df["date"] - pd.Timedelta(days=7)
    df["week"] = df["date"].dt.to_period("W").astype(str)

    return df


# -----------------------------
# APO OKUMA
# -----------------------------
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
    df["product_code"] = df["product_code"].astype(str)

    date_cols = df.columns[1:]

    df = df.melt(
        id_vars=["product_code"],
        value_vars=date_cols,
        var_name="date",
        value_name="outbound"
    )

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["outbound"] = pd.to_numeric(df["outbound"], errors="coerce").fillna(0)

    df = df.dropna(subset=["date"])
    df = df[df["outbound"] != 0]

    df["week"] = df["date"].dt.to_period("W").astype(str)

    return df


# -----------------------------
# UI
# -----------------------------
st.title("📦 Depo Kapasite Dashboard")

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

    try:
        supply = read_supply_file(supply_file)
    except Exception as e:
        st.error(f"Supply hata: {e}")
        st.stop()

    try:
        apo = read_apo_file(apo_file)
    except Exception as e:
        st.error(f"APO hata: {e}")
        st.stop()

    df = pd.merge(
        supply,
        apo,
        on=["product_code", "week"],
        how="outer"
    ).fillna(0)

    weekly = df.groupby("week")[["inbound", "outbound"]].sum().reset_index()

    weekly["stock"] = weekly["inbound"].cumsum() - weekly["outbound"].cumsum()

    weekly["palet"] = (weekly["stock"] / 2400).round(0)

    weekly["kapasite_%"] = (weekly["palet"] / depo_kapasite) * 100

    st.dataframe(
        weekly.style.format({
            "inbound": "{:,.0f}",
            "outbound": "{:,.0f}",
            "stock": "{:,.0f}",
            "palet": "{:,.0f}",
            "kapasite_%": "{:,.0f}"
        })
    )

    # ---------------- SAVE
    output = BytesIO()
    weekly.to_excel(output, index=False)

    st.session_state["report"] = output.getvalue()

# ---------------- SAVE PANEL
st.divider()
st.subheader("Kaydet")

if "report" in st.session_state:

    name = st.text_input("Rapor adı", "depo_rapor")

    if st.button("Geçmişe Kaydet"):

        filename = f"{name}.xlsx"
        path = HISTORY_DIR / filename

        path.write_bytes(st.session_state["report"])

        st.success("Kaydedildi")

# ---------------- HISTORY
st.divider()
st.subheader("Geçmiş")

files = list(HISTORY_DIR.glob("*.xlsx"))

if files:

    selected = st.selectbox("Seç", files)

    if st.button("Aç"):
        df = pd.read_excel(selected)
        st.dataframe(df)
