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

    # 🔍 TỰ ĐỘNG DÒ CỘT STORAGE LOCATION
    sloc_candidates = [
        c for c in df.columns
        if "storage" in c.lower() and "location" in c.lower()
    ]
    if not sloc_candidates:
        st.error("❌ Không tìm thấy cột Storage Location trong MB52")
        st.stop()

    sloc_col = sloc_candidates[0]
    df.rename(columns={sloc_col: "Storage Location"}, inplace=True)

    df["Unrestricted"] = pd.to_numeric(
        df["Unrestricted"], errors="coerce"
    ).fillna(0)

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
# MAP TỒN KHO 5 TẦNG
# =====================================================
map_da_cn = mb52_raw.groupby(
    ["Material", "Plant", "Storage Location", "WBS Element"],
    as_index=False
)["Unrestricted"].sum().set_index(
    ["Material", "Plant", "Storage Location", "WBS Element"]
)["Unrestricted"].to_dict()

map_da_tinh = mb52_raw.groupby(
    ["Material", "Plant", "WBS Element"],
    as_index=False
)["Unrestricted"].sum().set_index(
    ["Material", "Plant", "WBS Element"]
)["Unrestricted"].to_dict()

map_cn = mb52_raw.groupby(
    ["Material", "Plant", "Storage Location"],
    as_index=False
)["Unrestricted"].sum().set_index(
    ["Material", "Plant", "Storage Location"]
)["Unrestricted"].to_dict()

map_tinh = mb52_raw.groupby(
    ["Material", "Plant"],
    as_index=False
)["Unrestricted"].sum().set_index(
    ["Material", "Plant"]
)["Unrestricted"].to_dict()

map_kv = mb52_raw.groupby(
    ["Material"],
    as_index=False
)["Unrestricted"].sum().set_index(
    ["Material"]
)["Unrestricted"].to_dict()

# =====================================================
# UPLOAD PHIẾU XUẤT
# =====================================================
st.markdown("### 📂 Upload file phiếu xuất kho")

issue_file = st.file_uploader(
    "Upload file phiếu xuất kho", type=["xlsx", "xls"]
)
if not issue_file:
    st.stop()

@st.cache_data(show_spinner="🔄 Đang đọc file phiếu...")
def load_issue(file):
    df = pd.read_excel(file)
    df["Transfer Quantity"] = pd.to_numeric(
        df["Transfer Quantity"], errors="coerce"
    ).fillna(0)
    df["Actual Quantity"] = pd.to_numeric(
        df.get("Actual Quantity", 0), errors="coerce"
    ).fillna(0)
    return df

issue_df = load_issue(issue_file)

# =====================================================
# SIDEBAR – LỌC REALTIME (GIỮ NGUYÊN)
# =====================================================
st.sidebar.markdown("---")
st.sidebar.header("🔍 LỌC REALTIME")

filter_material = st.sidebar.text_input("Mã vật tư")

st.sidebar.markdown("**Functional Location**")
fl_search = st.sidebar.text_input("🔍 Tìm nhanh FL")

all_fl = sorted(issue_df["Functional Location"].dropna().unique())
if fl_search:
    all_fl = [f for f in all_fl if fl_search.lower() in str(f).lower()]

filter_fl = st.sidebar.multiselect("Chọn FL", all_fl)

filter_plant = st.sidebar.multiselect(
    "Plant", sorted(issue_df["Plant"].dropna().unique())
)

filter_status = st.sidebar.multiselect(
    "Tình trạng xuất kho",
    ["ĐẢM BẢO", "KHÔNG ĐẢM BẢO", "XUẤT ĐỦ", "KHÔNG ĐỦ"]
)

# =====================================================
# LUỸ KẾ 5 TẦNG
# =====================================================
def build_sequential_5_layer(df):
    r = df.copy()

    remain_da_cn = map_da_cn.copy()
    remain_da_tinh = map_da_tinh.copy()
    remain_cn = map_cn.copy()
    remain_tinh = map_tinh.copy()
    

    r["Tầng đáp ứng"] = ""
    r["Gợi ý chuyển WBS"] = ""
    r["Report Status"] = ""
    r["Thiếu kho"] = False

    for idx, row in r.iterrows():
        qty = row["Transfer Quantity"]
        mat = row["Material Number"]
        plant = row["Plant"]
        sloc = row["Sending Sloc"]
        wbs = row["Source WBS"]

        r.at[idx, "Tồn kho DA CN"] = remain_da_cn.get((mat, plant, sloc, wbs), 0)
        r.at[idx, "Tồn kho DA Tỉnh"] = remain_da_tinh.get((mat, plant, wbs), 0)
        r.at[idx, "Tồn kho CN"] = remain_cn.get((mat, plant, sloc), 0)
        r.at[idx, "Tồn kho Tỉnh"] = remain_tinh.get((mat, plant), 0)
        r.at[idx, "Tồn kho Khu vực"] = map_kv.get(mat, 0)

        layers = [
            ("Kho DA CN", remain_da_cn, (mat, plant, sloc, wbs)),
            ("Kho DA Tỉnh", remain_da_tinh, (mat, plant, wbs)),
            ("Kho CN", remain_cn, (mat, plant, sloc)),
            ("Kho Tỉnh", remain_tinh, (mat, plant))
            
        ]

        allocated = False
        for name, store, key in layers:
            cur = store.get(key, 0)
            if qty <= cur:
                store[key] = cur - qty
                allocated = True
                r.at[idx, "Tầng đáp ứng"] = name
                r.at[idx, "Report Status"] = "ĐẢM BẢO"
                if name != "Kho DA CN":
                    r.at[idx, "Gợi ý chuyển WBS"] = f"🧠 Có thể chuyển từ {name}"
                break

        if not allocated:
            kv_qty = map_kv.get(mat, 0)

            if qty <= kv_qty:
                r.at[idx, "Tầng đáp ứng"] = "Kho Khu vực (tham chiếu)"
                r.at[idx, "Report Status"] = "ĐẢM BẢO"
                r.at[idx, "Gợi ý chuyển WBS"] = "🧠 Có thể điều chuyển từ kho khu vực"
            else:
                r.at[idx, "Report Status"] = "KHÔNG ĐẢM BẢO"
                r.at[idx, "Thiếu kho"] = True
                r.at[idx, "Gợi ý chuyển WBS"] = "🚚 Thiếu toàn bộ các tầng kho"


    return r

# =====================================================
# APPLY FILTER
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

# =====================================================
# BUILD REPORT
# =====================================================
sequential_report = apply_filter(build_sequential_5_layer(issue_df))

# =====================================================
# DISPLAY
# =====================================================
cols = [
    "Request Number",
    "Material Number",
    "Material Description",
    "Plant",
    "Source WBS",
    "Sending Sloc",
    "Functional Location",
    "Transfer Quantity",
    "Tồn kho DA CN",
    "Tồn kho DA Tỉnh",
    "Tồn kho CN",
    "Tồn kho Tỉnh",
    "Tồn kho Khu vực",
    "Tầng đáp ứng",
    "Report Status",
    "Gợi ý chuyển WBS"
]

st.subheader("📊 BÁO CÁO KIỂM TRA")

st.dataframe(sequential_report[cols], use_container_width=True)

# =====================================================
# TỔNG HỢP THIẾU KHO THEO FL
# =====================================================
st.markdown("### 📊 TỔNG HỢP THIẾU KHO THEO FUNCTIONAL LOCATION")

summary = sequential_report[
    sequential_report["Thiếu kho"]
].groupby(
    "Functional Location"
).size().reset_index(
    name="Số dòng thiếu kho"
)

st.dataframe(summary, use_container_width=True)
