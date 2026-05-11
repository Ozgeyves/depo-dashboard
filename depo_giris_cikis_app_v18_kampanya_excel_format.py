import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="Depo Giriş Çıkış Raporu", layout="wide")

HISTORY_DIR = Path("history")
HISTORY_DIR.mkdir(exist_ok=True)

PASSWORDS = {
    "tedarik": "tedarik",
    "depo": "depo",
}

# =========================================================
# LOGIN
# =========================================================
def login_screen():
    st.title("📦 Depo Giriş Çıkış Raporu")
    password = st.text_input("Şifre", type="password")
    if st.button("Giriş Yap", type="primary"):
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
    text = str(x).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def normalize_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def clean_number(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, str):
        value = value.strip().replace(" ", "")
        # 1.234,56 formatını destekler
        if "," in value and "." in value:
            value = value.replace(".", "").replace(",", ".")
        elif "," in value:
            value = value.replace(",", ".")
    num = pd.to_numeric(value, errors="coerce")
    return 0.0 if pd.isna(num) else float(num)


def find_header_row_with_dates(raw: pd.DataFrame, max_rows: int = 20) -> int:
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


def week_label(date_series):
    return pd.to_datetime(date_series, errors="coerce", dayfirst=True).dt.to_period("W-MON").astype(str)


def pick_column(df: pd.DataFrame, keywords):
    cols = list(df.columns)
    lower_map = {str(c).lower().strip(): c for c in cols}
    for keyword in keywords:
        keyword = keyword.lower()
        for lower, original in lower_map.items():
            if keyword in lower:
                return original
    return None


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
        return pd.DataFrame(columns=["product_code", "date", "inbound_qty", "week", "source"])

    raw = pd.read_excel(file, sheet_name=0, header=None)
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
    melted["inbound_qty"] = melted["inbound_qty"].apply(clean_number)
    melted = melted.dropna(subset=["date"])
    melted = melted[melted["inbound_qty"] != 0]

    # Eski çalışan mantık korunuyor: supply tarihi 7 gün geri çekiliyor.
    melted["date"] = melted["date"] - pd.Timedelta(days=7)
    melted["week"] = week_label(melted["date"])
    melted["source"] = "Supply"
    return melted[["product_code", "date", "inbound_qty", "week", "source"]]


def read_apo_file(file):
    if file is None:
        return pd.DataFrame(columns=["product_code", "date", "outbound_qty", "week", "source"])

    raw = pd.read_excel(file, sheet_name=0, header=None)
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
    melted["outbound_qty"] = melted["outbound_qty"].apply(clean_number)
    melted = melted.dropna(subset=["date"])
    melted = melted[melted["outbound_qty"] != 0]
    melted["week"] = week_label(melted["date"])
    melted["source"] = "APO"
    return melted[["product_code", "date", "outbound_qty", "week", "source"]]


def read_product_type_file(file):
    if file is None:
        return pd.DataFrame(columns=["product_code", "product_type"])

    df = pd.read_excel(file, sheet_name=0)
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    code_col = pick_column(df, ["product_code", "product code", "ürün kod", "urun kod", "kod", "code", "product"])
    type_col = pick_column(df, ["ürün tipi", "urun tipi", "product_type", "product type", "tip", "type", "kategori", "category"])

    if code_col is None:
        code_col = first_non_empty_column(df)
    if type_col is None:
        possible = [c for c in df.columns if c != code_col]
        type_col = possible[0] if possible else code_col

    out = df[[code_col, type_col]].copy()
    out.columns = ["product_code", "product_type"]
    out["product_code"] = out["product_code"].apply(normalize_code)
    out["product_type"] = out["product_type"].apply(normalize_text)
    out = out[(out["product_code"] != "")]
    out = out.drop_duplicates(subset=["product_code"], keep="first")
    return out


def read_depo_data_file(file):
    if file is None:
        return pd.DataFrame()

    frames = []
    xls = pd.ExcelFile(file)
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(file, sheet_name=sheet)
            df = df.dropna(how="all")
            if df.empty:
                continue
            df.columns = [str(c).strip() for c in df.columns]
            df["sheet_name"] = sheet
            frames.append(df)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def manual_table_to_long(manual_df, qty_col_name, source_name):
    if manual_df is None or manual_df.empty:
        return pd.DataFrame(columns=["product_code", "date", qty_col_name, "week", "source"])

    df = manual_df.copy()
    df["product_code"] = df["product_code"].apply(normalize_code)
    df["date"] = pd.to_datetime(df["date"], errors="coerce", dayfirst=True)
    df[qty_col_name] = df[qty_col_name].apply(clean_number)
    df = df[(df["product_code"] != "") & df["date"].notna() & (df[qty_col_name] != 0)]
    df["week"] = week_label(df["date"])
    df["source"] = source_name
    return df[["product_code", "date", qty_col_name, "week", "source"]]

# =========================================================
# CALCULATION
# =========================================================
def classify_product_type(product_type):
    t = str(product_type).lower()
    if "mini" in t or "sample" in t or "numune" in t:
        return "Mini Sample"
    if "adr" in t:
        return "ADR"
    if t.strip() == "" or t in ["nan", "none"]:
        return "Ana Ürün"
    return product_type


def get_pallet_inner(product_type, ana_palet, mini_palet, adr_palet):
    t = str(product_type).lower()
    if "mini" in t or "sample" in t or "numune" in t:
        return mini_palet
    if "adr" in t:
        return adr_palet
    return ana_palet


def build_detail(supply, apo, manual_in, manual_out, product_types, ana_palet, mini_palet, adr_palet):
    inbound = pd.concat([supply, manual_in], ignore_index=True)
    outbound = pd.concat([apo, manual_out], ignore_index=True)

    inbound_group = inbound.groupby(["product_code", "week"], as_index=False)["inbound_qty"].sum()
    outbound_group = outbound.groupby(["product_code", "week"], as_index=False)["outbound_qty"].sum()

    detail = pd.merge(inbound_group, outbound_group, on=["product_code", "week"], how="outer")
    if detail.empty:
        return pd.DataFrame(columns=["product_code", "product_type", "week", "inbound_qty", "outbound_qty", "net_qty", "pallet_inner", "net_pallet"])

    detail["inbound_qty"] = pd.to_numeric(detail["inbound_qty"], errors="coerce").fillna(0)
    detail["outbound_qty"] = pd.to_numeric(detail["outbound_qty"], errors="coerce").fillna(0)
    detail["net_qty"] = detail["inbound_qty"] - detail["outbound_qty"]

    if product_types is not None and not product_types.empty:
        detail = detail.merge(product_types, on="product_code", how="left")
    else:
        detail["product_type"] = "Ana Ürün"

    detail["product_type"] = detail["product_type"].fillna("Ana Ürün").apply(classify_product_type)
    detail["pallet_inner"] = detail["product_type"].apply(lambda x: get_pallet_inner(x, ana_palet, mini_palet, adr_palet))
    detail["inbound_pallet"] = np.where(detail["pallet_inner"] > 0, detail["inbound_qty"] / detail["pallet_inner"], 0)
    detail["outbound_pallet"] = np.where(detail["pallet_inner"] > 0, detail["outbound_qty"] / detail["pallet_inner"], 0)
    detail["net_pallet"] = detail["inbound_pallet"] - detail["outbound_pallet"]
    return detail


def build_weekly(detail, starting_pallets, weekly_sarf_pallet, capacity_pallet, tir_kapasite):
    if detail.empty:
        return pd.DataFrame(columns=["week", "opening_pallet", "inbound_qty", "outbound_qty", "net_qty", "inbound_pallet", "outbound_pallet", "net_pallet", "weekly_sarf_pallet", "closing_pallet"])

    weekly = detail.groupby("week", as_index=False).agg(
        inbound_qty=("inbound_qty", "sum"),
        outbound_qty=("outbound_qty", "sum"),
        net_qty=("net_qty", "sum"),
        inbound_pallet=("inbound_pallet", "sum"),
        outbound_pallet=("outbound_pallet", "sum"),
        net_pallet=("net_pallet", "sum"),
    ).sort_values("week")

    weekly["weekly_sarf_pallet"] = weekly_sarf_pallet
    weekly["net_after_sarf_pallet"] = weekly["net_pallet"] - weekly["weekly_sarf_pallet"]
    weekly["opening_pallet"] = starting_pallets + weekly["net_after_sarf_pallet"].cumsum().shift(1).fillna(0)
    weekly["closing_pallet"] = starting_pallets + weekly["net_after_sarf_pallet"].cumsum()
    weekly["capacity_pallet"] = capacity_pallet
    weekly["empty_capacity_pallet"] = capacity_pallet - weekly["closing_pallet"]
    weekly["capacity_usage_%"] = np.where(capacity_pallet > 0, weekly["closing_pallet"] / capacity_pallet, 0)
    weekly["tir"] = np.where(tir_kapasite > 0, weekly["closing_pallet"] / tir_kapasite, 0)
    return weekly

# =========================================================
# UI
# =========================================================
st.title("📦 Depo Giriş - Çıkış / Stok Seviye Raporu")
logout_button()

with st.sidebar:
    st.header("Depo Kapasite Girdileri")
    baslangic_ana_palet = st.number_input("Başlangıç Ana Ürün Palet", value=0.0, step=1.0)
    baslangic_mini_palet = st.number_input("Başlangıç Mini Sample Palet", value=0.0, step=1.0)
    baslangic_adr_palet = st.number_input("Başlangıç ADR Palet", value=0.0, step=1.0)
    kapasite_palet = st.number_input("Toplam Depo Kapasitesi Palet", value=0.0, step=1.0)
    sarf = st.number_input("Haftalık Sarf Palet", value=250.0, step=1.0)

    if IS_TEDARIK:
        st.divider()
        st.subheader("Ürün Tipi Palet İçi")
        ana_palet = st.number_input("Ana Ürün Palet İçi", value=2400.0, step=1.0)
        mini_palet = st.number_input("Mini Sample Palet İçi", value=15000.0, step=1.0)
        adr_palet = st.number_input("ADR Palet İçi", value=5540.0, step=1.0)
        tir_kapasite = st.number_input("1 Tır Kaç Palet", value=40.0, step=1.0)
    else:
        ana_palet = 2400.0
        mini_palet = 15000.0
        adr_palet = 5540.0
        tir_kapasite = 40.0

st.subheader("Dosya Yükleme")
col1, col2, col3, col4 = st.columns(4)
with col1:
    supply_file = st.file_uploader("Supply dosyası", type=["xlsx", "xls", "xlsm"], key="supply")
with col2:
    apo_file = st.file_uploader("APO dosyası", type=["xlsx", "xls", "xlsm"], key="apo")
with col3:
    product_type_file = st.file_uploader("Ürün tipi dosyası", type=["xlsx", "xls", "xlsm"], key="product_type")
with col4:
    depo_file = st.file_uploader("Depo data dosyası", type=["xlsx", "xls", "xlsm"], key="depo_data")

st.subheader("Manuel Giriş / Çıkış")
manual_template = pd.DataFrame({
    "product_code": [""],
    "date": [pd.NaT],
    "qty": [0],
    "note": [""],
})

mcol1, mcol2 = st.columns(2)
with mcol1:
    st.caption("Ek girişleri buraya yazabilirsiniz.")
    manual_in_editor = st.data_editor(
        manual_template,
        num_rows="dynamic",
        use_container_width=True,
        key="manual_in_editor",
        column_config={"date": st.column_config.DateColumn("date", format="DD.MM.YYYY")},
    )
with mcol2:
    st.caption("Ek çıkışları buraya yazabilirsiniz.")
    manual_out_editor = st.data_editor(
        manual_template,
        num_rows="dynamic",
        use_container_width=True,
        key="manual_out_editor",
        column_config={"date": st.column_config.DateColumn("date", format="DD.MM.YYYY")},
    )

hesapla = st.button("Hesapla", type="primary")

if hesapla:
    try:
        if supply_file is None and apo_file is None and manual_in_editor.empty and manual_out_editor.empty:
            st.warning("Lütfen en az bir giriş/çıkış verisi yükleyin veya manuel satır ekleyin.")
            st.stop()

        supply = read_supply_file(supply_file)
        apo = read_apo_file(apo_file)
        product_types = read_product_type_file(product_type_file)
        depo_data = read_depo_data_file(depo_file)

        manual_in = manual_table_to_long(
            manual_in_editor.rename(columns={"qty": "inbound_qty"}),
            "inbound_qty",
            "Manuel Giriş",
        )
        manual_out = manual_table_to_long(
            manual_out_editor.rename(columns={"qty": "outbound_qty"}),
            "outbound_qty",
            "Manuel Çıkış",
        )

        starting_pallets = baslangic_ana_palet + baslangic_mini_palet + baslangic_adr_palet

        detail = build_detail(
            supply=supply,
            apo=apo,
            manual_in=manual_in,
            manual_out=manual_out,
            product_types=product_types,
            ana_palet=ana_palet,
            mini_palet=mini_palet,
            adr_palet=adr_palet,
        )

        weekly = build_weekly(
            detail=detail,
            starting_pallets=starting_pallets,
            weekly_sarf_pallet=sarf,
            capacity_pallet=kapasite_palet,
            tir_kapasite=tir_kapasite,
        )

        st.success("Rapor oluşturuldu.")

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Toplam Giriş Adet", f"{weekly['inbound_qty'].sum():,.0f}" if not weekly.empty else "0")
        kpi2.metric("Toplam Çıkış Adet", f"{weekly['outbound_qty'].sum():,.0f}" if not weekly.empty else "0")
        kpi3.metric("Son Stok Palet", f"{weekly['closing_pallet'].iloc[-1]:,.0f}" if not weekly.empty else f"{starting_pallets:,.0f}")
        if kapasite_palet > 0 and not weekly.empty:
            kpi4.metric("Son Kapasite Kullanımı", f"{weekly['capacity_usage_%'].iloc[-1]:.1%}")
        else:
            kpi4.metric("Hafta Sayısı", f"{weekly['week'].nunique():,.0f}" if not weekly.empty else "0")

        st.subheader("Haftalık Depo Özet")
        depo_weekly_cols = [
            "week", "opening_pallet", "inbound_qty", "outbound_qty", "net_qty",
            "weekly_sarf_pallet", "closing_pallet", "capacity_pallet",
            "empty_capacity_pallet", "capacity_usage_%"
        ]
        tedarik_extra_cols = ["inbound_pallet", "outbound_pallet", "net_pallet", "net_after_sarf_pallet", "tir"]

        if IS_TEDARIK:
            weekly_show_cols = [c for c in depo_weekly_cols + tedarik_extra_cols if c in weekly.columns]
        else:
            weekly_show_cols = [c for c in depo_weekly_cols if c in weekly.columns]

        st.dataframe(
            weekly[weekly_show_cols].style.format({
                "opening_pallet": "{:,.0f}",
                "inbound_qty": "{:,.0f}",
                "outbound_qty": "{:,.0f}",
                "net_qty": "{:,.0f}",
                "weekly_sarf_pallet": "{:,.0f}",
                "closing_pallet": "{:,.0f}",
                "capacity_pallet": "{:,.0f}",
                "empty_capacity_pallet": "{:,.0f}",
                "capacity_usage_%": "{:.1%}",
                "inbound_pallet": "{:,.1f}",
                "outbound_pallet": "{:,.1f}",
                "net_pallet": "{:,.1f}",
                "net_after_sarf_pallet": "{:,.1f}",
                "tir": "{:,.1f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Ürün Bazlı Detay")
        depo_detail_cols = ["product_code", "product_type", "week", "inbound_qty", "outbound_qty", "net_qty"]
        tedarik_detail_extra = ["pallet_inner", "inbound_pallet", "outbound_pallet", "net_pallet"]
        if IS_TEDARIK:
            detail_show_cols = [c for c in depo_detail_cols + tedarik_detail_extra if c in detail.columns]
        else:
            detail_show_cols = [c for c in depo_detail_cols if c in detail.columns]

        detail_sorted = detail.sort_values(["week", "product_code"]) if not detail.empty else detail
        st.dataframe(
            detail_sorted[detail_show_cols].style.format({
                "inbound_qty": "{:,.0f}",
                "outbound_qty": "{:,.0f}",
                "net_qty": "{:,.0f}",
                "pallet_inner": "{:,.0f}",
                "inbound_pallet": "{:,.1f}",
                "outbound_pallet": "{:,.1f}",
                "net_pallet": "{:,.1f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        if not product_types.empty:
            st.subheader("Ürün Tipi Kontrol")
            st.dataframe(product_types, use_container_width=True, hide_index=True)

        if depo_file is not None:
            st.subheader("Depo Data Bilgi Alanı")
            if depo_data.empty:
                st.info("Depo data dosyasında gösterilecek dolu veri bulunamadı.")
            else:
                st.dataframe(depo_data, use_container_width=True, hide_index=True)

        sheets = {
            "Haftalik Ozet": weekly[weekly_show_cols],
            "Urun Detay": detail_sorted[detail_show_cols],
        }
        if not product_types.empty:
            sheets["Urun Tipleri"] = product_types
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
            save_history(weekly[weekly_show_cols], detail_sorted[detail_show_cols])
        except Exception:
            pass

    except Exception as e:
        st.error("Rapor oluşturulurken hata oluştu.")
        st.exception(e)
else:
    st.info("Dosyaları yükleyip veya manuel giriş/çıkış ekleyip Hesapla butonuna basın.")
