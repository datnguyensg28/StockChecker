import streamlit as st
import pandas as pd
import os
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Stock Checker",
    layout="wide"
)

st.title("📦 PHẦN MỀM KIỂM TRA KHẢ NĂNG XUẤT KHO")

# =====================================================
# LOAD MB52 (CACHE – RẤT QUAN TRỌNG CHO FILE LỚN)
# =====================================================
@st.cache_data(show_spinner="🔄 Đang đọc MB52...")
def load_mb52():
    path = os.path.join(os.getcwd(), "MB52.xlsx")
    if not os.path.exists(path):
        st.error("❌ Không tìm thấy MB52.xlsx")
        st.stop()

    df = pd.read_excel(path)
    df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0)

    return (
        df.groupby(["Material", "Plant", "WBS Element"], as_index=False)["Unrestricted"]
        .sum()
    )

mb52_df = load_mb52()

# =====================================================
# UI – UPLOAD FILE
# =====================================================
uploaded_file = st.file_uploader(
    "📂 Upload file phiếu xuất kho",
    type=["xlsx", "xls"]
)

if not uploaded_file:
    st.stop()

@st.cache_data(show_spinner="🔄 Đang đọc file phiếu...")
def load_issue(file):
    df = pd.read_excel(file)
    df["Transfer Quantity"] = pd.to_numeric(df["Transfer Quantity"], errors="coerce").fillna(0)
    df["Actual Quantity"] = pd.to_numeric(df.get("Actual Quantity", 0), errors="coerce").fillna(0)
    return df

issue_df = load_issue(uploaded_file)

# =====================================================
# SIDEBAR – TÙY CHỌN
# =====================================================
st.sidebar.header("⚙️ TÙY CHỌN")

use_sequential = st.sidebar.checkbox(
    "Bật LUỸ KẾ TỒN KHO",
    value=False
)

sort_option = st.sidebar.selectbox(
    "Sắp xếp phiếu theo",
    ["Request Number", "Ngày phiếu", "Mức ưu tiên"],
    disabled=not use_sequential
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 LỌC NHANH")

filter_material = st.sidebar.text_input("Mã vật tư")
filter_wbs = st.sidebar.text_input("Source WBS")
filter_plant = st.sidebar.text_input("Plant")

# =====================================================
# DATA PREP
# =====================================================
stock_map = mb52_df.set_index(
    ["Material", "Plant", "WBS Element"]
)["Unrestricted"].to_dict()

# =====================================================
# SORT FUNCTION
# =====================================================
def sort_pending(df, option):
    if option == "Ngày phiếu" and "Request Date" in df.columns:
        return df.sort_values(by=["Request Date", "Request Number"])
    if option == "Mức ưu tiên" and "Priority" in df.columns:
        return df.sort_values(by=["Priority", "Request Number"])
    return df.sort_values(by=["Request Number"])

# =====================================================
# SIMPLE MODE
# =====================================================
def build_simple_report(df):
    result = df.copy()

    result["Tồn kho ban đầu"] = result.apply(
        lambda r: stock_map.get(
            (r["Material Number"], r["Plant"], r["Source WBS"]), 0
        ),
        axis=1
    )

    result["Tồn kho còn lại"] = result["Tồn kho ban đầu"]
    result["Âm tồn"] = ""

    def status(r):
        if r["Status"] == 12:
            return "XUẤT ĐỦ" if r["Transfer Quantity"] == r["Actual Quantity"] else "KHÔNG ĐỦ"
        return "ĐẢM BẢO" if r["Transfer Quantity"] <= r["Tồn kho ban đầu"] else "KHÔNG ĐẢM BẢO"

    result["Report Status"] = result.apply(status, axis=1)
    return result

# =====================================================
# SEQUENTIAL MODE (NÂNG CAO)
# =====================================================
def build_sequential_report(df):
    result = df.copy()
    pending = df[df["Status"].isin([1, 5, 9])].copy()
    pending = sort_pending(pending, sort_option)

    init_stock = stock_map.copy()
    remaining = stock_map.copy()

    result["Tồn kho ban đầu"] = 0
    result["Tồn kho còn lại"] = 0
    result["Âm tồn"] = ""

    for idx, row in pending.iterrows():
        key = (row["Material Number"], row["Plant"], row["Source WBS"])
        init_qty = init_stock.get(key, 0)
        remain = remaining.get(key, 0)

        if row["Transfer Quantity"] <= remain:
            result.at[idx, "Report Status"] = "ĐẢM BẢO"
            remaining[key] = remain - row["Transfer Quantity"]
        else:
            result.at[idx, "Report Status"] = "KHÔNG ĐẢM BẢO"
            result.at[idx, "Âm tồn"] = "⚠️"

        result.at[idx, "Tồn kho ban đầu"] = init_qty
        result.at[idx, "Tồn kho còn lại"] = remaining.get(key, remain)

    result.loc[result["Status"] == 12, "Report Status"] = result.apply(
        lambda r: "XUẤT ĐỦ" if r["Transfer Quantity"] == r["Actual Quantity"] else "KHÔNG ĐỦ",
        axis=1
    )

    return result

# =====================================================
# BUILD REPORTS
# =====================================================
simple_report = build_simple_report(issue_df)
sequential_report = build_sequential_report(issue_df)

# =====================================================
# FILTER FUNCTION (REALTIM)
# =====================================================
def apply_filter(df):
    if filter_material:
        df = df[df["Material Number"].astype(str).str.contains(filter_material)]
    if filter_wbs:
        df = df[df["Source WBS"].astype(str).str.contains(filter_wbs)]
    if filter_plant:
        df = df[df["Plant"].astype(str).str.contains(filter_plant)]
    return df

simple_report = apply_filter(simple_report)
sequential_report = apply_filter(sequential_report)

# =====================================================
# DISPLAY
# =====================================================
display_cols = [
    "Request Number",
    "Material Number",
    "Material Description",
    "Plant",
    "Source WBS",
    "Target WBS",
    "Target WBS Name",
    "Functional Location",
    "Sending Sloc",
    "Requirement Quantity",
    "Transfer Quantity",
    "Tồn kho ban đầu",
    "Tồn kho còn lại",
    "Report Status",
    "Âm tồn"
]

st.subheader("📊 BÁO CÁO KIỂM TRA")

tab1, tab2 = st.tabs(["📄 TÍNH THƯỜNG", "📊 LUỸ KẾ"])

with tab1:
    st.dataframe(simple_report[display_cols], use_container_width=True)

with tab2:
    st.dataframe(sequential_report[display_cols], use_container_width=True)

# =====================================================
# EXPORT
# =====================================================
def export_excel(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active

    color_map = {
        "KHÔNG ĐỦ": "FFC7CE",
        "XUẤT ĐỦ": "C6EFCE",
        "ĐẢM BẢO": "BDD7EE",
        "KHÔNG ĐẢM BẢO": "FFEB9C"
    }

    col = df.columns.get_loc("Report Status") + 1
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v in color_map:
            ws.cell(r, col).fill = PatternFill(
                start_color=color_map[v],
                end_color=color_map[v],
                fill_type="solid"
            )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.download_button(
        "📥 Export báo cáo TÍNH THƯỜNG",
        export_excel(simple_report),
        "BAO_CAO_TINH_THUONG.xlsx"
    )

with col2:
    st.download_button(
        "📥 Export báo cáo LUỸ KẾ",
        export_excel(sequential_report),
        "BAO_CAO_LUY_KE.xlsx"
    )
