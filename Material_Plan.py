import streamlit as st
import pandas as pd
import io

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="SAP MINI – Xuất kho BTS",
    layout="wide"
)

st.title("📦 SAP MINI – PHẦN MỀM VIẾT PHIẾU XUẤT KHO BTS")

# =====================================================
# UPLOAD FILES
# =====================================================
st.sidebar.header("📂 DỮ LIỆU ĐẦU VÀO")

mb52_file = st.sidebar.file_uploader("Upload MB52.xlsx", type=["xlsx"])
calloff_file = st.sidebar.file_uploader("Upload file Calloff", type=["xlsx"])

if not mb52_file or not calloff_file:
    st.stop()

# =====================================================
# OPTIONS
# =====================================================
st.sidebar.header("⚙️ TUỲ CHỌN TÍNH TOÁN")
use_sequential = st.sidebar.checkbox("🔁 TÍNH LUỸ KẾ TỒN KHO")

# =====================================================
# LOAD MB52
# =====================================================
@st.cache_data
def load_mb52(file):
    df = pd.read_excel(file)
    df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0)
    return df

mb52 = load_mb52(mb52_file)

# Tạo các tầng kho
stock_project = mb52.groupby(
    ["Material", "Plant", "Storage location", "WBS Element"],
    as_index=False
)["Unrestricted"].sum()

stock_sloc = mb52.groupby(
    ["Material", "Plant", "Storage location"],
    as_index=False
)["Unrestricted"].sum()

stock_plant = mb52.groupby(
    ["Material", "Plant"],
    as_index=False
)["Unrestricted"].sum()

stock_total = mb52.groupby(
    ["Material"],
    as_index=False
)["Unrestricted"].sum()

# Map tồn kho
map_project = stock_project.set_index(
    ["Material", "Plant", "Storage location", "WBS Element"]
)["Unrestricted"].to_dict()

map_sloc = stock_sloc.set_index(
    ["Material", "Plant", "Storage location"]
)["Unrestricted"].to_dict()

map_plant = stock_plant.set_index(
    ["Material", "Plant"]
)["Unrestricted"].to_dict()

map_total = stock_total.set_index(
    ["Material"]
)["Unrestricted"].to_dict()

# =====================================================
# LOAD CALLOFF
# =====================================================
calloff = pd.read_excel(calloff_file, sheet_name="calloff")
sm = pd.read_excel(calloff_file, sheet_name="SM")

# =====================================================
# BUILD REQUIREMENT
# =====================================================
records = []

material_cols = calloff.columns[4:]

for _, row in calloff.iterrows():
    ma_tram = row["ma_tram"]
    wbs_list = str(row["WBS Element"]).split(";")
    plant_priority = str(row["Plant"]).split(";")
    sloc_priority = str(row["Storage location"]).split(";")

    for mat_col in material_cols:
        qty_design = row[mat_col]
        if pd.isna(qty_design) or qty_design == 0:
            continue

        sm_rows = sm[sm["chung_loai_vat_tu"] == mat_col]

        for _, sm_row in sm_rows.iterrows():
            qty = qty_design * sm_row["dinh_muc"]

            records.append({
                "ma_tram": ma_tram,
                "chung_loai_vat_tu": sm_row["chung_loai_vat_tu"],
                "loai_vat_tu": sm_row["loai_vat_tu"],
                "Material": sm_row["Material"],
                "Material Description": sm_row["Material Description"],
                "Qty": qty,
                "Plant_Priority": plant_priority,
                "Sloc_Priority": sloc_priority,
                "WBS_List": wbs_list
            })

req_df = pd.DataFrame(records)

# =====================================================
# ENGINE XUẤT KHO
# =====================================================
results = []

remain_project = map_project.copy()
remain_sloc = map_sloc.copy()
remain_plant = map_plant.copy()
remain_total = map_total.copy()

for _, r in req_df.iterrows():
    material = r["Material"]
    qty = r["Qty"]
    status = "KHÔNG ĐỦ"
    note = ""

    for plant in r["Plant_Priority"]:
        for sloc in r["Sloc_Priority"]:
            for wbs in r["WBS_List"]:
                key_pj = (material, plant, sloc, wbs)
                stock = remain_project.get(key_pj, 0)

                if stock >= qty:
                    status = "ĐẢM BẢO"
                    note = "Xuất kho dự án"
                    if use_sequential:
                        remain_project[key_pj] -= qty
                    break
            if status == "ĐẢM BẢO":
                break

            key_sloc = (material, plant, sloc)
            stock = remain_sloc.get(key_sloc, 0)
            if stock >= qty:
                status = "ĐẢM BẢO (CHUYỂN WBS)"
                note = f"Chuyển WBS trong kho {sloc}"
                if use_sequential:
                    remain_sloc[key_sloc] -= qty
                break

        if status.startswith("ĐẢM BẢO"):
            break

        key_plant = (material, plant)
        stock = remain_plant.get(key_plant, 0)
        if stock >= qty:
            status = "KHÔNG ĐẢM BẢO"
            note = f"Gợi ý chuyển từ kho {plant}"
            break

    if status == "KHÔNG ĐỦ":
        if remain_total.get(material, 0) >= qty:
            note = "Có tồn kho tổng – cần điều phối"
        else:
            note = "Thiếu kho toàn hệ thống"

    results.append({
        **r,
        "Trạng thái": status,
        "Gợi ý": note
    })

report = pd.DataFrame(results)

# =====================================================
# FILTER REALTIME
# =====================================================
st.sidebar.header("🔍 LỌC REALTIME")

flt_material = st.sidebar.multiselect(
    "Material",
    sorted(report["Material"].unique())
)

flt_status = st.sidebar.multiselect(
    "Trạng thái",
    sorted(report["Trạng thái"].unique())
)

if flt_material:
    report = report[report["Material"].isin(flt_material)]

if flt_status:
    report = report[report["Trạng thái"].isin(flt_status)]

# =====================================================
# DISPLAY
# =====================================================
st.subheader("📊 KẾT QUẢ TÍNH TOÁN XUẤT KHO")

show_cols = [
    "ma_tram",
    "Material",
    "Material Description",
    "Qty",
    "Trạng thái",
    "Gợi ý"
]

st.dataframe(report[show_cols], use_container_width=True)

# =====================================================
# EXPORT
# =====================================================
def export_excel(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

st.download_button(
    "📥 Export báo cáo xuất kho",
    export_excel(report),
    "BAO_CAO_XUAT_KHO_BTS.xlsx"
)
