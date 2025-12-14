import streamlit as st
import pandas as pd
import os
import io
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(page_title="Stock Checker", layout="wide")
st.title("📦 PHẦN MỀM KIỂM TRA KHẢ NĂNG XUẤT KHO")

# =====================================================
# SIDEBAR – NGUỒN MB52
# =====================================================
st.sidebar.header("📦 NGUỒN TỒN KHO (MB52)")

mb52_source = st.sidebar.radio(
    "Chọn nguồn dữ liệu tồn kho",
    ["☁️ MB52 mặc định (Datnd5 update)", "📂 Upload file MB52"]
)

def find_mb52_path():
    if not os.path.exists("data"):
        return None
    for f in os.listdir("data"):
        if f.lower() == "mb52.xlsx":
            return os.path.join("data", f)
    return None

@st.cache_data(show_spinner="🔄 Đang đọc MB52...")
def load_mb52(file):
    df = pd.read_excel(file)
    df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0)
    return df

if mb52_source == "📂 Upload file MB52":
    mb52_upload = st.sidebar.file_uploader("Upload MB52.xlsx", type=["xlsx"])
    if not mb52_upload:
        st.stop()
    mb52_raw = load_mb52(mb52_upload)
else:
    mb52_path = find_mb52_path()
    if not mb52_path:
        st.error("❌ Không tìm thấy MB52.xlsx trong thư mục data/")
        st.stop()
    mb52_raw = load_mb52(mb52_path)

    upload_time = (
        datetime.datetime.utcfromtimestamp(os.path.getmtime(mb52_path))
        + datetime.timedelta(hours=7)
    ).strftime("%d/%m/%Y %H:%M")

    st.info(f"ℹ️ Tồn kho tại thời điểm upload MB52 (giờ VN): **{upload_time}**")

# =====================================================
# MAP TỒN KHO
# =====================================================
stock_wbs = mb52_raw.groupby(
    ["Material", "Plant", "WBS Element"], as_index=False
)["Unrestricted"].sum()

stock_total = mb52_raw.groupby(
    ["Material", "Plant"], as_index=False
)["Unrestricted"].sum()

map_wbs = stock_wbs.set_index(
    ["Material", "Plant", "WBS Element"]
)["Unrestricted"].to_dict()

map_total = stock_total.set_index(
    ["Material", "Plant"]
)["Unrestricted"].to_dict()

# =====================================================
# UPLOAD PHIẾU XUẤT
# =====================================================
st.markdown("### 📂 Upload file phiếu xuất kho")

issue_file = st.file_uploader("Upload file phiếu xuất kho", type=["xlsx", "xls"])
if not issue_file:
    st.stop()

@st.cache_data(show_spinner="🔄 Đang đọc file phiếu...")
def load_issue(file):
    df = pd.read_excel(file)
    df["Transfer Quantity"] = pd.to_numeric(df["Transfer Quantity"], errors="coerce").fillna(0)
    df["Actual Quantity"] = pd.to_numeric(df.get("Actual Quantity", 0), errors="coerce").fillna(0)
    return df

issue_df = load_issue(issue_file)

# =====================================================
# SIDEBAR – TÙY CHỌN
# =====================================================
st.sidebar.header("⚙️ TUỲ CHỌN TÍNH TOÁN")
use_sequential = st.sidebar.checkbox("🔁 Bật LUỸ KẾ TỒN KHO")

sort_option = st.sidebar.selectbox(
    "Sắp xếp phiếu theo",
    ["Request Number", "Ngày phiếu", "Mức ưu tiên"],
    disabled=not use_sequential
)

# =====================================================
# SIDEBAR – LỌC REALTIME (CHECKBOX)
# =====================================================
st.sidebar.markdown("---")
st.sidebar.header("🔍 LỌC REALTIME")

filter_material = st.sidebar.text_input("Mã vật tư")

# Functional Location (checkbox + search)
st.sidebar.markdown("**Functional Location**")
fl_search = st.sidebar.text_input("🔍 Tìm nhanh FL")

all_fl = sorted(issue_df["Functional Location"].dropna().unique())
if fl_search:
    all_fl = [f for f in all_fl if fl_search.lower() in str(f).lower()]

filter_fl = st.sidebar.multiselect("Chọn FL", all_fl)

# Plant checkbox
filter_plant = st.sidebar.multiselect(
    "Plant",
    sorted(issue_df["Plant"].dropna().unique())
)

# Status checkbox
filter_status = st.sidebar.multiselect(
    "Tình trạng xuất kho",
    ["ĐẢM BẢO", "KHÔNG ĐẢM BẢO", "XUẤT ĐỦ", "KHÔNG ĐỦ"]
)

# =====================================================
# SORT
# =====================================================
def sort_pending(df):
    if sort_option == "Ngày phiếu" and "Request Date" in df.columns:
        return df.sort_values(["Request Date", "Request Number"])
    if sort_option == "Mức ưu tiên" and "Priority" in df.columns:
        return df.sort_values(["Priority", "Request Number"])
    return df.sort_values("Request Number")

# =====================================================
# TÍNH THƯỜNG
# =====================================================
def build_simple(df):
    r = df.copy()

    r["Tồn kho WBS"] = r.apply(
        lambda x: map_wbs.get(
            (x["Material Number"], x["Plant"], x["Source WBS"]), 0
        ), axis=1
    )

    r["Tồn kho tổng"] = r.apply(
        lambda x: map_total.get(
            (x["Material Number"], x["Plant"]), 0
        ), axis=1
    )

    def status(x):
        if x["Status"] == 12:
            return "XUẤT ĐỦ" if x["Transfer Quantity"] == x["Actual Quantity"] else "KHÔNG ĐỦ"
        return "ĐẢM BẢO" if x["Transfer Quantity"] <= x["Tồn kho WBS"] else "KHÔNG ĐẢM BẢO"

    r["Report Status"] = r.apply(status, axis=1)

    r["Gợi ý chuyển WBS"] = r.apply(
        lambda x:
        "🧠 Có thể chuyển WBS nội bộ"
        if x["Report Status"] == "KHÔNG ĐẢM BẢO"
        and x["Transfer Quantity"] <= x["Tồn kho tổng"]
        else "",
        axis=1
    )

    r["Thiếu kho"] = r["Report Status"] == "KHÔNG ĐẢM BẢO"
    return r

# =====================================================
# LUỸ KẾ
# =====================================================
def build_sequential(df):
    r = build_simple(df)
    pending = r[r["Status"].isin([1, 5, 9])].copy()
    pending = sort_pending(pending)

    remain = map_wbs.copy()

    for idx, row in pending.iterrows():
        key = (row["Material Number"], row["Plant"], row["Source WBS"])
        cur = remain.get(key, 0)

        if row["Transfer Quantity"] <= cur:
            remain[key] = cur - row["Transfer Quantity"]
        else:
            r.at[idx, "Thiếu kho"] = True

    return r

simple_report = build_simple(issue_df)
sequential_report = build_sequential(issue_df)

# =====================================================
# FILTER APPLY
# =====================================================
def apply_filter(df):
    if filter_material:
        df = df[df["Material Number"].astype(str).str.contains(filter_material)]
    if filter_fl:
        df = df[df["Functional Location"].isin(filter_fl)]
    if filter_plant:
        df = df[df["Plant"].isin(filter_plant)]
    if filter_status:
        df = df[df["Report Status"].isin(filter_status)]
    return df

simple_report = apply_filter(simple_report)
sequential_report = apply_filter(sequential_report)

# =====================================================
# DISPLAY
# =====================================================
cols = [
    "Request Number",
    "Material Number",
    "Material Description",
    "Plant",
    "Source WBS",
    "Functional Location",
    "Transfer Quantity",
    "Tồn kho WBS",
    "Tồn kho tổng",
    "Report Status",
    "Gợi ý chuyển WBS"
]

st.subheader("📊 BÁO CÁO KIỂM TRA")

tab1, tab2 = st.tabs(["📄 TÍNH THƯỜNG", "📊 LUỸ KẾ"])

with tab1:
    st.dataframe(simple_report[cols], use_container_width=True)

with tab2:
    st.dataframe(sequential_report[cols], use_container_width=True)

# =====================================================
# TỔNG HỢP THIẾU KHO THEO FL
# =====================================================
st.markdown("### 📊 TỔNG HỢP THIẾU KHO THEO FUNCTIONAL LOCATION")

summary = simple_report[simple_report["Thiếu kho"]].groupby(
    "Functional Location"
).size().reset_index(name="Số dòng thiếu kho")

st.dataframe(summary, use_container_width=True)
