# =====================================================
# STOCKFLOW CHECKER - KIEM TRA DAM BAO XUAT KHO
# Author: DatND5
# Version: 3.0 Streamlit
# =====================================================

import datetime
import io
from typing import Any, Dict, Optional, Tuple

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="StockFlow Checker",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =====================================================
# CONFIG
# =====================================================
DEFAULT_MB52_RAW_URL = "https://raw.githubusercontent.com/datnguyensg28/StockChecker/main/data/MB52.XLSX"
LOCAL_MB52_PATH = "data/MB52.XLSX"

APP_NAME = "StockFlow Checker"
APP_SUBTITLE = "Kiểm tra phiếu xuất kho theo trạng thái thực xuất và tồn kho MB52"
APP_VERSION = "3.0"

REQUIRED_MB52_COLUMNS = ["Material", "Plant", "Unrestricted", "WBS Element"]
REQUIRED_ISSUE_COLUMNS = [
    "Request Number",
    "Material Number",
    "Material Description",
    "Plant",
    "Source WBS",
    "Sending Sloc",
    "Functional Location",
    "Transfer Quantity",
]

DETAIL_COLUMNS = [
    "Request Number",
    "Material Number",
    "Material Description",
    "Plant",
    "Source WBS",
    "Sending Sloc",
    "Functional Location",
    "Transfer Quantity",
    "Actual Quantity",
    "Status",
    "Còn thiếu",
    "Tình trạng",
    "Gợi ý xử lý",
]

STOCK_COLUMNS = [
    "Tồn kho DA CN",
    "Tồn kho DA Tỉnh",
    "Tồn kho CN",
    "Tồn kho Tỉnh",
    "Tồn kho Khu vực",
]


# =====================================================
# CSS
# =====================================================
st.markdown(
    """
    <style>
        .main .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 1180px;
        }
        .app-header {
            padding: 18px 0 10px 0;
            border-bottom: 1px solid #e5e7eb;
            margin-bottom: 18px;
        }
        .app-title {
            font-size: 32px;
            font-weight: 800;
            color: #111827;
            line-height: 1.15;
        }
        .app-subtitle {
            color: #4b5563;
            margin-top: 6px;
            font-size: 15px;
        }
        .step-title {
            margin: 18px 0 10px 0;
            padding: 12px 14px;
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            border-left: 5px solid #2563eb;
            border-radius: 8px;
            font-weight: 800;
            color: #111827;
        }
        div[data-testid="stMetric"] {
            background: #ffffff;
            padding: 14px 16px;
            border-radius: 8px;
            border: 1px solid #e5e7eb;
        }
        div[data-testid="stMetricValue"] {
            font-size: 26px;
            font-weight: 800;
        }
        .result-card {
            border-radius: 8px;
            padding: 22px 24px;
            margin: 14px 0 16px 0;
            border: 1px solid;
        }
        .result-ok {
            background: #ecfdf5;
            border-color: #86efac;
            color: #14532d;
        }
        .result-bad {
            background: #fff7ed;
            border-color: #fdba74;
            color: #7c2d12;
        }
        .result-headline {
            font-size: 30px;
            font-weight: 900;
            margin-bottom: 8px;
        }
        .result-copy {
            font-size: 17px;
            font-weight: 600;
        }
        .small-note {
            color: #6b7280;
            font-size: 13px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================================================
# HELPERS
# =====================================================
def get_mb52_raw_url() -> str:
    try:
        return str(st.secrets.get("MB52_RAW_URL", DEFAULT_MB52_RAW_URL)).strip()
    except Exception:
        return DEFAULT_MB52_RAW_URL


def normalize_key_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_column_name(value: Any) -> str:
    return str(value).strip().lower()


def stop_with_missing_columns(missing: list[str], file_label: str) -> None:
    st.error(f"❌ File {file_label} thiếu cột bắt buộc: {', '.join(missing)}")
    st.stop()


def validate_columns(df: pd.DataFrame, required_cols: list[str], file_label: str) -> None:
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        stop_with_missing_columns(missing, file_label)


def detect_storage_location_column(df: pd.DataFrame) -> Optional[str]:
    exact_candidates = [c for c in df.columns if normalize_column_name(c) == "storage location"]
    if exact_candidates:
        return exact_candidates[0]

    fuzzy_candidates = [
        c
        for c in df.columns
        if "storage" in normalize_column_name(c) and "location" in normalize_column_name(c)
    ]
    if fuzzy_candidates:
        return fuzzy_candidates[0]
    return None


def detect_column_by_name_or_position(
    df: pd.DataFrame,
    accepted_names: list[str],
    excel_column_index: int,
    display_name: str,
) -> str:
    normalized_names = {name.strip().lower() for name in accepted_names}
    for col in df.columns:
        if normalize_column_name(col) in normalized_names:
            return col

    zero_based_index = excel_column_index - 1
    if len(df.columns) > zero_based_index:
        return df.columns[zero_based_index]

    st.error(
        f"❌ Không tìm thấy cột {display_name}. "
        f"Hãy đặt tên cột là {accepted_names[0]} hoặc đặt đúng vị trí cột Excel."
    )
    st.stop()


def normalize_status(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def is_exported_status(value: Any) -> bool:
    return normalize_status(value) == "12"


@st.cache_data(ttl=300, show_spinner="Đang tải MB52 mới nhất từ GitHub...")
def download_mb52_from_github(raw_url: str) -> Tuple[bytes, Dict[str, str]]:
    if not raw_url:
        raise ValueError("Chưa cấu hình GitHub Raw URL MB52.")

    headers = {
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "User-Agent": "StockFlow-Checker/3.0",
    }
    response = requests.get(raw_url, headers=headers, timeout=60)
    response.raise_for_status()

    meta = {
        "source": "GitHub - MB52 mới nhất",
        "url": raw_url,
        "loaded_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "last_modified": response.headers.get("Last-Modified", ""),
        "etag": response.headers.get("ETag", ""),
    }
    return response.content, meta


@st.cache_data(show_spinner="Đang đọc MB52 local...")
def read_local_mb52(path: str) -> Tuple[bytes, Dict[str, str]]:
    with open(path, "rb") as file:
        content = file.read()
    meta = {
        "source": "Local - data/MB52.XLSX",
        "url": path,
        "loaded_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "last_modified": "",
        "etag": "",
    }
    return content, meta


@st.cache_data(show_spinner="Đang đọc MB52...")
def load_mb52(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))

    sloc_col = detect_storage_location_column(df)
    if not sloc_col:
        st.error("❌ Không tìm thấy cột Storage Location trong MB52.")
        st.stop()
    if sloc_col != "Storage Location":
        df = df.rename(columns={sloc_col: "Storage Location"})

    validate_columns(df, REQUIRED_MB52_COLUMNS + ["Storage Location"], "MB52")

    df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0)
    for col in ["Material", "Plant", "Storage Location", "WBS Element"]:
        df[col] = df[col].apply(normalize_key_value)

    return df


@st.cache_data(show_spinner="Đang đọc file phiếu xuất kho...")
def load_issue(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))
    validate_columns(df, REQUIRED_ISSUE_COLUMNS, "phiếu xuất kho")

    actual_col = detect_column_by_name_or_position(
        df,
        ["Actual Quantity", "Thực xuất"],
        28,  # AB
        "Actual Quantity / Thực xuất",
    )
    status_col = detect_column_by_name_or_position(
        df,
        ["Status"],
        29,  # AC
        "Status",
    )

    if actual_col != "Actual Quantity":
        df = df.rename(columns={actual_col: "Actual Quantity"})
    if status_col != "Status":
        df = df.rename(columns={status_col: "Status"})

    df["Transfer Quantity"] = pd.to_numeric(df["Transfer Quantity"], errors="coerce").fillna(0)
    df["Actual Quantity"] = pd.to_numeric(df["Actual Quantity"], errors="coerce").fillna(0)
    df["Status"] = df["Status"].apply(normalize_status)

    for col in ["Material Number", "Plant", "Source WBS", "Sending Sloc", "Functional Location"]:
        df[col] = df[col].apply(normalize_key_value)

    return df


def build_inventory_maps(mb52_raw: pd.DataFrame):
    map_da_cn = (
        mb52_raw.groupby(["Material", "Plant", "Storage Location", "WBS Element"], as_index=False)["Unrestricted"]
        .sum()
        .set_index(["Material", "Plant", "Storage Location", "WBS Element"])["Unrestricted"]
        .to_dict()
    )
    map_da_tinh = (
        mb52_raw.groupby(["Material", "Plant", "WBS Element"], as_index=False)["Unrestricted"]
        .sum()
        .set_index(["Material", "Plant", "WBS Element"])["Unrestricted"]
        .to_dict()
    )
    map_cn = (
        mb52_raw.groupby(["Material", "Plant", "Storage Location"], as_index=False)["Unrestricted"]
        .sum()
        .set_index(["Material", "Plant", "Storage Location"])["Unrestricted"]
        .to_dict()
    )
    map_tinh = (
        mb52_raw.groupby(["Material", "Plant"], as_index=False)["Unrestricted"]
        .sum()
        .set_index(["Material", "Plant"])["Unrestricted"]
        .to_dict()
    )
    map_kv = mb52_raw.groupby(["Material"], as_index=False)["Unrestricted"].sum().set_index(["Material"])["Unrestricted"].to_dict()

    return map_da_cn, map_da_tinh, map_cn, map_tinh, map_kv


def build_sequential_5_layer(issue_df: pd.DataFrame, mb52_raw: pd.DataFrame) -> pd.DataFrame:
    map_da_cn, map_da_tinh, map_cn, map_tinh, map_kv = build_inventory_maps(mb52_raw)

    r = issue_df.copy()
    remain_da_cn = map_da_cn.copy()
    remain_da_tinh = map_da_tinh.copy()
    remain_cn = map_cn.copy()
    remain_tinh = map_tinh.copy()

    r["Tầng đáp ứng"] = ""
    r["Gợi ý chuyển WBS"] = ""
    r["Report Status"] = ""
    r["Thiếu kho"] = False

    for col in STOCK_COLUMNS:
        r[col] = 0.0

    for idx, row in r.iterrows():
        qty = float(row["Transfer Quantity"])
        mat = normalize_key_value(row["Material Number"])
        plant = normalize_key_value(row["Plant"])
        sloc = normalize_key_value(row["Sending Sloc"])
        wbs = normalize_key_value(row["Source WBS"])

        da_cn_key = (mat, plant, sloc, wbs)
        da_tinh_key = (mat, plant, wbs)
        cn_key = (mat, plant, sloc)
        tinh_key = (mat, plant)
        kv_key = mat

        r.at[idx, "Tồn kho DA CN"] = remain_da_cn.get(da_cn_key, 0)
        r.at[idx, "Tồn kho DA Tỉnh"] = remain_da_tinh.get(da_tinh_key, 0)
        r.at[idx, "Tồn kho CN"] = remain_cn.get(cn_key, 0)
        r.at[idx, "Tồn kho Tỉnh"] = remain_tinh.get(tinh_key, 0)
        r.at[idx, "Tồn kho Khu vực"] = map_kv.get(kv_key, 0)

        da_cn_qty = remain_da_cn.get(da_cn_key, 0)
        if qty <= da_cn_qty:
            remain_da_cn[da_cn_key] = da_cn_qty - qty
            r.at[idx, "Tầng đáp ứng"] = "Kho DA CN"
            r.at[idx, "Report Status"] = "ĐẢM BẢO"
            continue

        r.at[idx, "Report Status"] = "KHÔNG ĐẢM BẢO"
        r.at[idx, "Thiếu kho"] = True

        if qty <= remain_da_tinh.get(da_tinh_key, 0):
            remain_da_tinh[da_tinh_key] -= qty
            r.at[idx, "Gợi ý chuyển WBS"] = "Có thể chuyển từ Kho DA Tỉnh"
        elif qty <= remain_cn.get(cn_key, 0):
            remain_cn[cn_key] -= qty
            r.at[idx, "Gợi ý chuyển WBS"] = "Có thể chuyển từ Kho CN"
        elif qty <= remain_tinh.get(tinh_key, 0):
            remain_tinh[tinh_key] -= qty
            r.at[idx, "Gợi ý chuyển WBS"] = "Có thể chuyển từ Kho Tỉnh"
        elif qty <= map_kv.get(kv_key, 0):
            r.at[idx, "Gợi ý chuyển WBS"] = "Có thể điều chuyển từ Kho Khu vực"
        else:
            r.at[idx, "Gợi ý chuyển WBS"] = "Thiếu toàn bộ các tầng kho"

    return r


def build_business_conclusion(report_df: pd.DataFrame) -> pd.DataFrame:
    r = report_df.copy()

    exported = r["Status"].apply(is_exported_status)
    enough_actual = r["Actual Quantity"] >= r["Transfer Quantity"]
    enough_mb52 = ~r["Thiếu kho"]

    shortage_by_actual = (r["Transfer Quantity"] - r["Actual Quantity"]).clip(lower=0)
    shortage_by_mb52 = r["Transfer Quantity"].where(~enough_mb52, 0)
    r["Còn thiếu"] = shortage_by_actual.where(shortage_by_actual > 0, shortage_by_mb52).fillna(0)

    r["Tình trạng"] = "Đảm bảo xuất kho"
    r["Gợi ý xử lý"] = "Không cần xử lý thêm"

    not_exported_mask = ~exported
    short_actual_mask = exported & ~enough_actual
    mb52_missing_mask = exported & enough_actual & ~enough_mb52

    r.loc[not_exported_mask, "Tình trạng"] = "Chưa xuất kho"
    r.loc[not_exported_mask, "Gợi ý xử lý"] = "Kiểm tra trạng thái phiếu, thực hiện xuất kho để Status = 12"

    r.loc[short_actual_mask, "Tình trạng"] = "Xuất thiếu"
    r.loc[short_actual_mask, "Gợi ý xử lý"] = "Kiểm tra số lượng thực xuất và xuất bổ sung phần còn thiếu"

    r.loc[mb52_missing_mask, "Tình trạng"] = "Thiếu tồn kho MB52"
    r.loc[mb52_missing_mask, "Gợi ý xử lý"] = r.loc[mb52_missing_mask, "Gợi ý chuyển WBS"].fillna("")
    r.loc[mb52_missing_mask & (r["Gợi ý xử lý"] == ""), "Gợi ý xử lý"] = "Kiểm tra bổ sung tồn kho MB52 hoặc điều chuyển vật tư"

    r["Đảm bảo 100%"] = exported & enough_actual & enough_mb52
    return r


def build_conclusion_sheet(total: int, ok: int, not_ok: int, mb52_meta: Dict[str, str]) -> pd.DataFrame:
    ok_rate = (ok / total * 100) if total else 0
    conclusion = (
        "ĐẢM BẢO XUẤT KHO 100%"
        if total > 0 and not_ok == 0
        else "CHƯA ĐẢM BẢO XUẤT KHO 100%"
    )
    return pd.DataFrame(
        [
            {"Thông tin": "Kết luận", "Giá trị": conclusion},
            {"Thông tin": "Tổng dòng", "Giá trị": total},
            {"Thông tin": "Đã xuất đủ", "Giá trị": ok},
            {"Thông tin": "Chưa đảm bảo", "Giá trị": not_ok},
            {"Thông tin": "Tỷ lệ đảm bảo", "Giá trị": f"{ok_rate:.1f}%"},
            {"Thông tin": "Nguồn MB52", "Giá trị": mb52_meta.get("source", "")},
            {"Thông tin": "MB52 URL/Path", "Giá trị": mb52_meta.get("url", "")},
            {"Thông tin": "Thời điểm kiểm tra", "Giá trị": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
        ]
    )


def auto_width_worksheet(ws) -> None:
    for col_idx, column_cells in enumerate(ws.columns, 1):
        max_length = 0
        for cell in column_cells:
            cell_length = len(str(cell.value)) if cell.value is not None else 0
            max_length = max(max_length, cell_length)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_length + 2, 12), 48)


def format_workbook(writer, sheet_names: list[str]) -> None:
    wb = writer.book
    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)
    bad_fill = PatternFill("solid", fgColor="FFEDD5")

    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        if sheet_name == "ChiTietChuaDamBao":
            for row in range(2, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = bad_fill
        auto_width_worksheet(ws)


def export_excel(full_df: pd.DataFrame, issue_df: pd.DataFrame, mb52_meta: Dict[str, str]) -> bytes:
    total = len(full_df)
    ok = int(full_df["Đảm bảo 100%"].sum())
    not_ok = total - ok
    error_df = full_df.loc[~full_df["Đảm bảo 100%"], DETAIL_COLUMNS].copy()
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet_names = ["KetLuan"]
        build_conclusion_sheet(total, ok, not_ok, mb52_meta).to_excel(writer, index=False, sheet_name="KetLuan")

        if not_ok > 0:
            error_df.to_excel(writer, index=False, sheet_name="ChiTietChuaDamBao")
            error_df[
                [
                    "Request Number",
                    "Material Number",
                    "Material Description",
                    "Plant",
                    "Functional Location",
                    "Còn thiếu",
                    "Tình trạng",
                    "Gợi ý xử lý",
                ]
            ].to_excel(writer, index=False, sheet_name="GoiYXuLy")
            sheet_names.extend(["ChiTietChuaDamBao", "GoiYXuLy"])

        format_workbook(writer, sheet_names)

    return output.getvalue()


def render_result_card(is_all_ok: bool) -> None:
    if is_all_ok:
        st.markdown(
            """
            <div class="result-card result-ok">
                <div class="result-headline">✅ ĐẢM BẢO XUẤT KHO 100%</div>
                <div class="result-copy">Loại phiếu này đã xuất đủ toàn bộ vật tư. Không cần xử lý thêm.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="result-card result-bad">
                <div class="result-headline">⚠️ CHƯA ĐẢM BẢO XUẤT KHO 100%</div>
                <div class="result-copy">Chỉ các dòng lỗi hoặc chưa đủ được hiển thị bên dưới.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# =====================================================
# UI
# =====================================================
st.markdown(
    f"""
    <div class="app-header">
        <div class="app-title">📦 {APP_NAME}</div>
        <div class="app-subtitle">{APP_SUBTITLE} · Version {APP_VERSION}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="step-title">Bước 1: Chọn nguồn MB52</div>', unsafe_allow_html=True)
source_options = [
    "GitHub - MB52 mới nhất",
    "Local - data/MB52.XLSX",
    "Upload MB52 tạm thời",
]
mb52_source = st.radio("Nguồn dữ liệu MB52", source_options, horizontal=True, label_visibility="collapsed")

mb52_bytes: Optional[bytes] = None
mb52_meta: Dict[str, str] = {}

col_source, col_refresh = st.columns([4, 1])
with col_source:
    if mb52_source == "GitHub - MB52 mới nhất":
        raw_url = st.text_input("GitHub Raw URL MB52", value=get_mb52_raw_url())
        try:
            mb52_bytes, mb52_meta = download_mb52_from_github(raw_url)
        except Exception as exc:
            st.error(f"❌ Không tải được MB52 từ GitHub: {exc}")
            st.stop()
    elif mb52_source == "Local - data/MB52.XLSX":
        try:
            mb52_bytes, mb52_meta = read_local_mb52(LOCAL_MB52_PATH)
        except Exception as exc:
            st.error(f"❌ Không đọc được file local {LOCAL_MB52_PATH}: {exc}")
            st.stop()
    else:
        upload_mb52 = st.file_uploader("Upload MB52 tạm thời", type=["xlsx", "xls"], key="mb52_upload")
        if not upload_mb52:
            st.info("Vui lòng upload file MB52 để tiếp tục.")
            st.stop()
        mb52_bytes = upload_mb52.getvalue()
        mb52_meta = {
            "source": "Upload MB52 tạm thời",
            "url": upload_mb52.name,
            "loaded_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "last_modified": "",
            "etag": "",
        }

with col_refresh:
    st.write("")
    st.write("")
    if st.button("🔄 Làm mới MB52", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

mb52_raw = load_mb52(mb52_bytes)
st.success(
    f"Đã sẵn sàng MB52: {len(mb52_raw):,} dòng · "
    f"{mb52_raw['Material'].nunique():,} mã vật tư · "
    f"nguồn {mb52_meta.get('source', '')}"
)

st.markdown('<div class="step-title">Bước 2: Upload phiếu xuất kho</div>', unsafe_allow_html=True)
issue_file = st.file_uploader(
    "Chọn file phiếu xuất kho",
    type=["xlsx", "xls"],
    help="File cần có Transfer Quantity, Actual Quantity ở cột AB hoặc theo tên cột, và Status ở cột AC hoặc theo tên cột.",
)
if not issue_file:
    st.info("Upload phiếu xuất kho để phần mềm kết luận ngay.")
    st.stop()

issue_df = load_issue(issue_file.getvalue())

with st.spinner("Đang kiểm tra trạng thái thực xuất và tồn kho MB52 theo 5 tầng..."):
    stock_report = build_sequential_5_layer(issue_df, mb52_raw)
    final_report = build_business_conclusion(stock_report)

total_lines = len(final_report)
ok_lines = int(final_report["Đảm bảo 100%"].sum())
not_ok_lines = total_lines - ok_lines
ok_rate = (ok_lines / total_lines * 100) if total_lines else 0
is_all_ok = total_lines > 0 and not_ok_lines == 0

st.markdown('<div class="step-title">Bước 3: Xem kết luận</div>', unsafe_allow_html=True)
metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("Tổng dòng", f"{total_lines:,}")
metric2.metric("Đã xuất đủ", f"{ok_lines:,}")
metric3.metric("Chưa đảm bảo", f"{not_ok_lines:,}")
metric4.metric("Tỷ lệ đảm bảo", f"{ok_rate:.1f}%")

render_result_card(is_all_ok)

error_df = final_report.loc[~final_report["Đảm bảo 100%"], DETAIL_COLUMNS].copy()

if not is_all_ok:
    error_counts = error_df["Tình trạng"].value_counts().rename_axis("Tình trạng").reset_index(name="Số dòng")
    st.dataframe(error_counts, use_container_width=True, hide_index=True, height=150)

    st.dataframe(
        error_df,
        use_container_width=True,
        hide_index=True,
        height=430,
        column_config={
            "Transfer Quantity": st.column_config.NumberColumn("Transfer Quantity", format="%.2f"),
            "Actual Quantity": st.column_config.NumberColumn("Actual Quantity", format="%.2f"),
            "Còn thiếu": st.column_config.NumberColumn("Còn thiếu", format="%.2f"),
            "Gợi ý xử lý": st.column_config.TextColumn("Gợi ý xử lý", width="large"),
        },
    )

export_bytes = export_excel(final_report, issue_df, mb52_meta)
file_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
st.download_button(
    label="⬇️ Tải kết quả Excel",
    data=export_bytes,
    file_name=f"StockFlow_KetQua_XuatKho_{file_time}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.caption("StockFlow Checker · Người dùng upload phiếu, phần mềm trả lời ngay: đảm bảo 100% hoặc thiếu dòng nào, vì sao, xử lý thế nào.")
