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
st.set_page_config(
    page_title="Stock Checker",
    layout="wide"
)

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
    data_dir = "data"
    if not os.path.exists(data_dir):
        return None

    for f in os.listdir(data_dir):
        if f.lower() == "mb52.xlsx":
            return os.path.join(data_dir, f)
    return None

# =====================================================
# LOAD MB52
# =====================================================


@st.cache_data(show_spinner="🔄 Đang đọc MB52...")
def load_mb52_from_file(file):
    df = pd.read_excel(file)
    df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0)
    return df.groupby(
        ["Material", "Plant", "WBS Element"],
        as_index=False
    )["Unrestricted"].sum()


if mb52_source == "📂 Upload file MB52":
    mb52_upload = st.sidebar.file_uploader(
        "Upload MB52.xlsx",
        type=["xlsx"]
    )
    if not mb52_upload:
        st.warning("⚠️ Vui lòng upload file MB52")
        st.stop()

    mb52_df = load_mb52_from_file(mb52_upload)

else:
    mb52_path = find_mb52_path()
    if not mb52_path:
        st.error("❌ Không tìm thấy MB52.xlsx trong thư mục data/")
        st.stop()

    mb52_df = load_mb52_from_file(mb52_path)

    
    upload_time = (
    datetime.datetime.utcfromtimestamp(os.path.getmtime(mb52_path))
    + datetime.timedelta(hours=7)
    ).strftime("%d/%m/%Y %H:%M")

    st.info(
        f"ℹ️ **Lưu ý:** Tồn kho hiển thị được tính tại thời điểm "
        f"file MB52 upload lên server (giờ Việt Nam) vào lúc: "
        f"**{upload_time}**. "
        f"Dữ liệu không phản ánh tồn kho realtime."
    )

# =====================================================
# UPLOAD PHIẾU XUẤT
# =====================================================
st.markdown("### 📂 Upload file phiếu xuất kho")

issue_file = st.file_uploader(
    "Upload file phiếu xuất kho",
    type=["xlsx", "xls"]
)

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
# SIDEBAR – TUỲ CHỌN
# =====================================================
st.sidebar.header("⚙️ TUỲ CHỌN TÍNH TOÁN")

use_sequential = st.sidebar.checkbox("🔁 Bật LUỸ KẾ TỒN KHO")

sort_option = st.sidebar.selectbox(
    "Sắp xếp phiếu theo",
    ["Request Number", "Ngày phiếu", "Mức ưu tiên"],
    disabled=not use_sequential
)

st.sidebar.markdown("---")
st.sidebar.header("🔍 LỌC REALTIME")

filter_material = st.sidebar.text_input("Mã vật tư")
filter_wbs = st.sidebar.text_input("Source WBS")
filter_plant = st.sidebar.text_input("Plant")

# =====================================================
# MAP TỒN KHO
# =====================================================
stock_map = mb52_df.set_index(
    ["Material", "Plant", "WBS Element"]
)["Unrestricted"].to_dict()

# =====================================================
# SORT
# =====================================================
def sort_pending(df, option):
    if option == "Ngày phiếu" and "Request Date" in df.columns:
        return df.sort_values(["Request Date", "Request Number"])
    if option == "Mức ưu tiên" and "Priority" in df.columns:
        return df.sort_values(["Priority", "Request Number"])
    return df.sort_values("Request Number")

# =====================================================
# TÍNH THƯỜNG
# =====================================================
def build_simple_report(df):
    r = df.copy()

    r["Tồn kho ban đầu"] = r.apply(
        lambda x: stock_map.get(
            (x["Material Number"], x["Plant"], x["Source WBS"]), 0
        ),
        axis=1
    )
    r["Tồn kho còn lại"] = r["Tồn kho ban đầu"]
    r["Âm tồn"] = ""

    def status(x):
        if x["Status"] == 12:
            return "XUẤT ĐỦ" if x["Transfer Quantity"] == x["Actual Quantity"] else "KHÔNG ĐỦ"
        return "ĐẢM BẢO" if x["Transfer Quantity"] <= x["Tồn kho ban đầu"] else "KHÔNG ĐẢM BẢO"

    r["Report Status"] = r.apply(status, axis=1)
    return r

# =====================================================
# LUỸ KẾ
# =====================================================
def build_sequential_report(df):
    r = df.copy()
    pending = df[df["Status"].isin([1, 5, 9])].copy()
    pending = sort_pending(pending, sort_option)

    remaining = stock_map.copy()

    r["Tồn kho ban đầu"] = 0
    r["Tồn kho còn lại"] = 0
    r["Âm tồn"] = ""

    for idx, row in pending.iterrows():
        key = (row["Material Number"], row["Plant"], row["Source WBS"])
        init_qty = stock_map.get(key, 0)
        remain = remaining.get(key, 0)

        if row["Transfer Quantity"] <= remain:
            r.at[idx, "Report Status"] = "ĐẢM BẢO"
            remaining[key] = remain - row["Transfer Quantity"]
        else:
            r.at[idx, "Report Status"] = "KHÔNG ĐẢM BẢO"
            r.at[idx, "Âm tồn"] = "⚠️"

        r.at[idx, "Tồn kho ban đầu"] = init_qty
        r.at[idx, "Tồn kho còn lại"] = remaining.get(key, remain)

    r.loc[r["Status"] == 12, "Report Status"] = r.apply(
        lambda x: "XUẤT ĐỦ" if x["Transfer Quantity"] == x["Actual Quantity"] else "KHÔNG ĐỦ",
        axis=1
    )

    return r

# =====================================================
# BUILD
# =====================================================
simple_report = build_simple_report(issue_df)
sequential_report = build_sequential_report(issue_df)

# =====================================================
# FILTER
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
cols = [
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

tab1, tab2 = st.tabs(["📄 TÍNH TOÁN TỪNG PHIẾU XUẤT KHO", "📊 TÍNH THEO LUỸ KẾ"])

with tab1:
    st.dataframe(simple_report[cols], use_container_width=True)

with tab2:
    st.dataframe(sequential_report[cols], use_container_width=True)

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
c1, c2 = st.columns(2)

with c1:
    st.download_button(
        "📥 Export báo cáo TÍNH THƯỜNG",
        export_excel(simple_report),
        "BAO_CAO_TINH_THUONG.xlsx"
    )

with c2:
    st.download_button(
        "📥 Export báo cáo LUỸ KẾ",
        export_excel(sequential_report),
        "BAO_CAO_LUY_KE.xlsx"
    )
