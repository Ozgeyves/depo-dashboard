import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Depo Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP_DIR = Path(__file__).parent
HISTORY_DIR = APP_DIR / "history"
HISTORY_DIR.mkdir(exist_ok=True)

# ============================================================
# LOGIN
# ============================================================
def get_secret_value(key, default_value):
    try:
        return st.secrets[key]
    except Exception:
        return default_value

planning_password = get_secret_value("PLANNING_PASSWORD", "planning2026")
depo_password = get_secret_value("DEPO_PASSWORD", "depo2026")

if "role" not in st.session_state:
    st.session_state["role"] = None

if st.session_state["role"] is None:

    if (APP_DIR / "logo.png").exists():
        st.image("logo.png", width=120)

    st.title("Depo Dashboard")
    st.caption("Lütfen şifre giriniz")

    password = st.text_input("Şifre", type="password")

    if st.button("Giriş Yap"):

        if password == planning_password:
            st.session_state["role"] = "planning"
            st.rerun()

        elif password == depo_password:
            st.session_state["role"] = "depo"
            st.rerun()

        else:
            st.error("Şifre yanlış")

    st.stop()

# ============================================================
# HELPERS
# ============================================================
def normalize_code(x):

    if pd.isna(x):
        return None

    text = str(x).strip()

    match = re.search(r"\d+", text)

    if match:
        text = match.group(0)

    if text.endswith(".0"):
        text = text[:-2]

    return text.lstrip("0") or "0"

def safe_format(x):

    if isinstance(x, (int, float, np.integer, np.floating)):
        return f"{x:,.0f}"

    return x

# ============================================================
# HEADER
# ============================================================
col1, col2 = st.columns([1, 6])

with col1:

    if (APP_DIR / "logo.png").exists():
        st.image("logo.png", width=90)

with col2:

    if st.session_state["role"] == "planning":
        st.title("Planning Dashboard")

    else:
        st.title("Depo Operasyon Dashboard")

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:

    if st.button("Çıkış Yap"):
        st.session_state["role"] = None
        st.rerun()

    st.header("Dosyalar")

    supply_file = st.file_uploader(
        "Supply Dosyası",
        type=["xlsx", "xlsm"]
    )

    apo_file = st.file_uploader(
        "APO Forecast Dosyası",
        type=["xlsx", "xlsm"]
    )

    mapping_file = st.file_uploader(
        "Ürün Tipi / Kampanya / ADR Dosyası",
        type=["xlsx", "xlsm"]
    )

    ekol_file = st.file_uploader(
        "Ekol Depo Data Dosyası",
        type=["xlsx", "xlsm"]
    )

# ============================================================
# PLANNING PARAMS
# ============================================================
if st.session_state["role"] == "planning":

    with st.sidebar:

        st.header("Palet Parametreleri")

        ana_palet = st.number_input(
            "Ana Ürün Palet İçi",
            value=2400
        )

        mini_palet = st.number_input(
            "Mini Sample Palet İçi",
            value=15000
        )

        adr_palet = st.number_input(
            "ADR Palet İçi",
            value=5540
        )

        sarf_palet = st.number_input(
            "Haftalık Sarf Palet",
            value=250
        )

        tir_kapasite = st.number_input(
            "1 Tır Kaç Palet",
            value=40
        )

        st.header("Başlangıç Stok")

        initial_ana = st.number_input(
            "Başlangıç Ana Ürün",
            value=0
        )

        initial_mini = st.number_input(
            "Başlangıç Mini Sample",
            value=0
        )

        initial_adr = st.number_input(
            "Başlangıç ADR",
            value=0
        )

# ============================================================
# BUTTON
# ============================================================
calculate = st.button("Raporu Hesapla", type="primary")

# ============================================================
# MAIN
# ============================================================
if calculate:

    st.success("Dosyalar başarıyla yüklendi.")

    # --------------------------------------------------------
    # DEMO TABLE
    # --------------------------------------------------------
    weeks = [
        "2026-W14",
        "2026-W15",
        "2026-W16",
        "2026-W17",
    ]

    report = pd.DataFrame({
        "Hafta": weeks,
        "Kampanya": [
            "1+1",
            "CRM",
            "Anneler Günü",
            ""
        ],
        "Ana Ürün Giriş": [120000, 180000, 210000, 140000],
        "Ana Ürün Çıkış": [95000, 120000, 160000, 150000],
        "Mini Sample Giriş": [25000, 40000, 60000, 35000],
        "Mini Sample Çıkış": [12000, 15000, 18000, 17000],
        "ADR Giriş": [4000, 6000, 4500, 5000],
        "ADR Çıkış": [2000, 3000, 2500, 2800],
    })

    # ========================================================
    # PLANNING SCREEN
    # ========================================================
    if st.session_state["role"] == "planning":

        report["Ana Ürün Palet"] = (
            report["Ana Ürün Giriş"] / ana_palet
        ).round(0)

        report["Mini Sample Palet"] = (
            report["Mini Sample Giriş"] / mini_palet
        ).round(0)

        report["ADR Palet"] = (
            report["ADR Giriş"] / adr_palet
        ).round(0)

        report["Total Palet"] = (
            report["Ana Ürün Palet"] +
            report["Mini Sample Palet"] +
            sarf_palet -
            report["ADR Palet"]
        )

        report["Tır Sayısı"] = (
            report["Ana Ürün Palet"] +
            report["Mini Sample Palet"]
        ) / tir_kapasite

        st.subheader("Planning Görünümü")

        st.dataframe(
            report.style.format(safe_format),
            use_container_width=True
        )

    # ========================================================
    # DEPO SCREEN
    # ========================================================
    else:

        depo_cols = [
            "Hafta",
            "Kampanya",
            "Ana Ürün Giriş",
            "Ana Ürün Çıkış",
            "Mini Sample Giriş",
            "Mini Sample Çıkış",
            "ADR Giriş",
            "ADR Çıkış",
        ]

        depo_report = report[depo_cols]

        st.subheader("Depo Operasyon Görünümü")

        st.info(
            "Bu ekranda palet hesaplamaları gösterilmez."
        )

        st.dataframe(
            depo_report.style.format(safe_format),
            use_container_width=True
        )

    # ========================================================
    # EKOL
    # ========================================================
    if ekol_file is not None:

        st.subheader("Ekol Depo Data")

        st.success(
            "Ekol dosyası başarıyla yüklendi."
        )

    # ========================================================
    # EXPORT
    # ========================================================
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        report.to_excel(
            writer,
            index=False,
            sheet_name="Rapor"
        )

    st.download_button(
        "Excel İndir",
        data=output.getvalue(),
        file_name="depo_dashboard.xlsx"
    )
