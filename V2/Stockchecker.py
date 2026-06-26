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
    page_icon="ðŸ“¦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =====================================================
# CONFIG
# =====================================================
DEFAULT_MB52_RAW_URL = "https://raw.githubusercontent.com/datnguyensg28/StockChecker/main/data/MB52.XLSX"
LOCAL_MB52_PATH = "data/MB52.XLSX"

APP_NAME = "StockFlow Checker"
APP_SUBTITLE = "Kiá»ƒm tra phiáº¿u xuáº¥t kho theo tráº¡ng thÃ¡i thá»±c xuáº¥t vÃ  tá»“n kho MB52"
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
    "CÃ²n thiáº¿u",
    "TÃ¬nh tráº¡ng",
    "Gá»£i Ã½ xá»­ lÃ½",
]

STOCK_COLUMNS = [
    "Tá»“n kho DA CN",
    "Tá»“n kho DA Tá»‰nh",
    "Tá»“n kho CN",
    "Tá»“n kho Tá»‰nh",
    "Tá»“n kho Khu vá»±c",
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
    st.error(f"âŒ File {file_label} thiáº¿u cá»™t báº¯t buá»™c: {', '.join(missing)}")
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
        f"âŒ KhÃ´ng tÃ¬m tháº¥y cá»™t {display_name}. "
        f"HÃ£y Ä‘áº·t tÃªn cá»™t lÃ  {accepted_names[0]} hoáº·c Ä‘áº·t Ä‘Ãºng vá»‹ trÃ­ cá»™t Excel."
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


@st.cache_data(ttl=300, show_spinner="Äang táº£i MB52 má»›i nháº¥t tá»« GitHub...")
def download_mb52_from_github(raw_url: str) -> Tuple[bytes, Dict[str, str]]:
    if not raw_url:
        raise ValueError("ChÆ°a cáº¥u hÃ¬nh GitHub Raw URL MB52.")

    headers = {
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "User-Agent": "StockFlow-Checker/3.0",
    }
    response = requests.get(raw_url, headers=headers, timeout=60)
    response.raise_for_status()

    meta = {
        "source": "GitHub - MB52 má»›i nháº¥t",
        "url": raw_url,
        "loaded_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "last_modified": response.headers.get("Last-Modified", ""),
        "etag": response.headers.get("ETag", ""),
    }
    return response.content, meta


@st.cache_data(show_spinner="Äang Ä‘á»c MB52 local...")
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


@st.cache_data(show_spinner="Äang Ä‘á»c MB52...")
def load_mb52(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))

    sloc_col = detect_storage_location_column(df)
    if not sloc_col:
        st.error("âŒ KhÃ´ng tÃ¬m tháº¥y cá»™t Storage Location trong MB52.")
        st.stop()
    if sloc_col != "Storage Location":
        df = df.rename(columns={sloc_col: "Storage Location"})

    validate_columns(df, REQUIRED_MB52_COLUMNS + ["Storage Location"], "MB52")

    df["Unrestricted"] = pd.to_numeric(df["Unrestricted"], errors="coerce").fillna(0)
    for col in ["Material", "Plant", "Storage Location", "WBS Element"]:
        df[col] = df[col].apply(normalize_key_value)

    return df


@st.cache_data(show_spinner="Äang Ä‘á»c file phiáº¿u xuáº¥t kho...")
def load_issue(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))
    validate_columns(df, REQUIRED_ISSUE_COLUMNS, "phiáº¿u xuáº¥t kho")

    actual_col = detect_column_by_name_or_position(
        df,
        ["Actual Quantity", "Thá»±c xuáº¥t"],
        28,  # AB
        "Actual Quantity / Thá»±c xuáº¥t",
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

    r["Táº§ng Ä‘Ã¡p á»©ng"] = ""
    r["Gá»£i Ã½ chuyá»ƒn WBS"] = ""
    r["Report Status"] = ""
    r["Thiáº¿u kho"] = False

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

        r.at[idx, "Tá»“n kho DA CN"] = remain_da_cn.get(da_cn_key, 0)
        r.at[idx, "Tá»“n kho DA Tá»‰nh"] = remain_da_tinh.get(da_tinh_key, 0)
        r.at[idx, "Tá»“n kho CN"] = remain_cn.get(cn_key, 0)
        r.at[idx, "Tá»“n kho Tá»‰nh"] = remain_tinh.get(tinh_key, 0)
        r.at[idx, "Tá»“n kho Khu vá»±c"] = map_kv.get(kv_key, 0)

        da_cn_qty = remain_da_cn.get(da_cn_key, 0)
        if qty <= da_cn_qty:
            remain_da_cn[da_cn_key] = da_cn_qty - qty
            r.at[idx, "Táº§ng Ä‘Ã¡p á»©ng"] = "Kho DA CN"
            r.at[idx, "Report Status"] = "Äáº¢M Báº¢O"
            continue

        r.at[idx, "Report Status"] = "KHÃ”NG Äáº¢M Báº¢O"
        r.at[idx, "Thiáº¿u kho"] = True

        if qty <= remain_da_tinh.get(da_tinh_key, 0):
            remain_da_tinh[da_tinh_key] -= qty
            r.at[idx, "Gá»£i Ã½ chuyá»ƒn WBS"] = "CÃ³ thá»ƒ chuyá»ƒn tá»« Kho DA Tá»‰nh"
        elif qty <= remain_cn.get(cn_key, 0):
            remain_cn[cn_key] -= qty
            r.at[idx, "Gá»£i Ã½ chuyá»ƒn WBS"] = "CÃ³ thá»ƒ chuyá»ƒn tá»« Kho CN"
        elif qty <= remain_tinh.get(tinh_key, 0):
            remain_tinh[tinh_key] -= qty
            r.at[idx, "Gá»£i Ã½ chuyá»ƒn WBS"] = "CÃ³ thá»ƒ chuyá»ƒn tá»« Kho Tá»‰nh"
        elif qty <= map_kv.get(kv_key, 0):
            r.at[idx, "Gá»£i Ã½ chuyá»ƒn WBS"] = "CÃ³ thá»ƒ Ä‘iá»u chuyá»ƒn tá»« Kho Khu vá»±c"
        else:
            r.at[idx, "Gá»£i Ã½ chuyá»ƒn WBS"] = "Thiáº¿u toÃ n bá»™ cÃ¡c táº§ng kho"

    return r


def build_business_conclusion(report_df: pd.DataFrame) -> pd.DataFrame:
    r = report_df.copy()

    exported = r["Status"].apply(is_exported_status)
    enough_actual = r["Actual Quantity"] >= r["Transfer Quantity"]
    enough_mb52 = ~r["Thiáº¿u kho"]

    shortage_by_actual = (r["Transfer Quantity"] - r["Actual Quantity"]).clip(lower=0)
    shortage_by_mb52 = r["Transfer Quantity"].where(~enough_mb52, 0)
    r["CÃ²n thiáº¿u"] = shortage_by_actual.where(shortage_by_actual > 0, shortage_by_mb52).fillna(0)

    r["TÃ¬nh tráº¡ng"] = "Äáº£m báº£o xuáº¥t kho"
    r["Gá»£i Ã½ xá»­ lÃ½"] = "KhÃ´ng cáº§n xá»­ lÃ½ thÃªm"

    not_exported_mask = ~exported
    short_actual_mask = exported & ~enough_actual
    mb52_missing_mask = exported & enough_actual & ~enough_mb52

    r.loc[not_exported_mask, "TÃ¬nh tráº¡ng"] = "ChÆ°a xuáº¥t kho"
    r.loc[not_exported_mask, "Gá»£i Ã½ xá»­ lÃ½"] = "Kiá»ƒm tra tráº¡ng thÃ¡i phiáº¿u, thá»±c hiá»‡n xuáº¥t kho Ä‘á»ƒ Status = 12"

    r.loc[short_actual_mask, "TÃ¬nh tráº¡ng"] = "Xuáº¥t thiáº¿u"
    r.loc[short_actual_mask, "Gá»£i Ã½ xá»­ lÃ½"] = "Kiá»ƒm tra sá»‘ lÆ°á»£ng thá»±c xuáº¥t vÃ  xuáº¥t bá»• sung pháº§n cÃ²n thiáº¿u"

    r.loc[mb52_missing_mask, "TÃ¬nh tráº¡ng"] = "Thiáº¿u tá»“n kho MB52"
    r.loc[mb52_missing_mask, "Gá»£i Ã½ xá»­ lÃ½"] = r.loc[mb52_missing_mask, "Gá»£i Ã½ chuyá»ƒn WBS"].fillna("")
    r.loc[mb52_missing_mask & (r["Gá»£i Ã½ xá»­ lÃ½"] == ""), "Gá»£i Ã½ xá»­ lÃ½"] = "Kiá»ƒm tra bá»• sung tá»“n kho MB52 hoáº·c Ä‘iá»u chuyá»ƒn váº­t tÆ°"

    r["Äáº£m báº£o 100%"] = exported & enough_actual & enough_mb52
    return r


def build_conclusion_sheet(total: int, ok: int, not_ok: int, mb52_meta: Dict[str, str]) -> pd.DataFrame:
    ok_rate = (ok / total * 100) if total else 0
    conclusion = (
        "Äáº¢M Báº¢O XUáº¤T KHO 100%"
        if total > 0 and not_ok == 0
        else "CHÆ¯A Äáº¢M Báº¢O XUáº¤T KHO 100%"
    )
    return pd.DataFrame(
        [
            {"ThÃ´ng tin": "Káº¿t luáº­n", "GiÃ¡ trá»‹": conclusion},
            {"ThÃ´ng tin": "Tá»•ng dÃ²ng", "GiÃ¡ trá»‹": total},
            {"ThÃ´ng tin": "ÄÃ£ xuáº¥t Ä‘á»§", "GiÃ¡ trá»‹": ok},
            {"ThÃ´ng tin": "ChÆ°a Ä‘áº£m báº£o", "GiÃ¡ trá»‹": not_ok},
            {"ThÃ´ng tin": "Tá»· lá»‡ Ä‘áº£m báº£o", "GiÃ¡ trá»‹": f"{ok_rate:.1f}%"},
            {"ThÃ´ng tin": "Nguá»“n MB52", "GiÃ¡ trá»‹": mb52_meta.get("source", "")},
            {"ThÃ´ng tin": "MB52 URL/Path", "GiÃ¡ trá»‹": mb52_meta.get("url", "")},
            {"ThÃ´ng tin": "Thá»i Ä‘iá»ƒm kiá»ƒm tra", "GiÃ¡ trá»‹": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
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
    ok = int(full_df["Äáº£m báº£o 100%"].sum())
    not_ok = total - ok
    error_df = full_df.loc[~full_df["Äáº£m báº£o 100%"], DETAIL_COLUMNS].copy()
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
                    "CÃ²n thiáº¿u",
                    "TÃ¬nh tráº¡ng",
                    "Gá»£i Ã½ xá»­ lÃ½",
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
                <div class="result-headline">âœ… Äáº¢M Báº¢O XUáº¤T KHO 100%</div>
                <div class="result-copy">Loáº¡i phiáº¿u nÃ y Ä‘Ã£ xuáº¥t Ä‘á»§ toÃ n bá»™ váº­t tÆ°. KhÃ´ng cáº§n xá»­ lÃ½ thÃªm.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="result-card result-bad">
                <div class="result-headline">âš ï¸ CHÆ¯A Äáº¢M Báº¢O XUáº¤T KHO 100%</div>
                <div class="result-copy">Chá»‰ cÃ¡c dÃ²ng lá»—i hoáº·c chÆ°a Ä‘á»§ Ä‘Æ°á»£c hiá»ƒn thá»‹ bÃªn dÆ°á»›i.</div>
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
        <div class="app-title">ðŸ“¦ {APP_NAME}</div>
        <div class="app-subtitle">{APP_SUBTITLE} Â· Version {APP_VERSION}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="step-title">BÆ°á»›c 1: Chá»n nguá»“n MB52</div>', unsafe_allow_html=True)
source_options = [
    "GitHub - MB52 má»›i nháº¥t",
    "Local - data/MB52.XLSX",
    "Upload MB52 táº¡m thá»i",
]
mb52_source = st.radio("Nguá»“n dá»¯ liá»‡u MB52", source_options, horizontal=True, label_visibility="collapsed")

mb52_bytes: Optional[bytes] = None
mb52_meta: Dict[str, str] = {}

col_source, col_refresh = st.columns([4, 1])
with col_source:
    if mb52_source == "GitHub - MB52 má»›i nháº¥t":
        raw_url = st.text_input("GitHub Raw URL MB52", value=get_mb52_raw_url())
        try:
            mb52_bytes, mb52_meta = download_mb52_from_github(raw_url)
        except Exception as exc:
            st.error(f"âŒ KhÃ´ng táº£i Ä‘Æ°á»£c MB52 tá»« GitHub: {exc}")
            st.stop()
    elif mb52_source == "Local - data/MB52.XLSX":
        try:
            mb52_bytes, mb52_meta = read_local_mb52(LOCAL_MB52_PATH)
        except Exception as exc:
            st.error(f"âŒ KhÃ´ng Ä‘á»c Ä‘Æ°á»£c file local {LOCAL_MB52_PATH}: {exc}")
            st.stop()
    else:
        upload_mb52 = st.file_uploader("Upload MB52 táº¡m thá»i", type=["xlsx", "xls"], key="mb52_upload")
        if not upload_mb52:
            st.info("Vui lÃ²ng upload file MB52 Ä‘á»ƒ tiáº¿p tá»¥c.")
            st.stop()
        mb52_bytes = upload_mb52.getvalue()
        mb52_meta = {
            "source": "Upload MB52 táº¡m thá»i",
            "url": upload_mb52.name,
            "loaded_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "last_modified": "",
            "etag": "",
        }

with col_refresh:
    st.write("")
    st.write("")
    if st.button("ðŸ”„ LÃ m má»›i MB52", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

mb52_raw = load_mb52(mb52_bytes)
st.success(
    f"ÄÃ£ sáºµn sÃ ng MB52: {len(mb52_raw):,} dÃ²ng Â· "
    f"{mb52_raw['Material'].nunique():,} mÃ£ váº­t tÆ° Â· "
    f"nguá»“n {mb52_meta.get('source', '')}"
)

st.markdown('<div class="step-title">BÆ°á»›c 2: Upload phiáº¿u xuáº¥t kho</div>', unsafe_allow_html=True)
issue_file = st.file_uploader(
    "Chá»n file phiáº¿u xuáº¥t kho",
    type=["xlsx", "xls"],
    help="File cáº§n cÃ³ Transfer Quantity, Actual Quantity á»Ÿ cá»™t AB hoáº·c theo tÃªn cá»™t, vÃ  Status á»Ÿ cá»™t AC hoáº·c theo tÃªn cá»™t.",
)
if not issue_file:
    st.info("Upload phiáº¿u xuáº¥t kho Ä‘á»ƒ pháº§n má»m káº¿t luáº­n ngay.")
    st.stop()

issue_df = load_issue(issue_file.getvalue())

with st.spinner("Äang kiá»ƒm tra tráº¡ng thÃ¡i thá»±c xuáº¥t vÃ  tá»“n kho MB52 theo 5 táº§ng..."):
    stock_report = build_sequential_5_layer(issue_df, mb52_raw)
    final_report = build_business_conclusion(stock_report)

total_lines = len(final_report)
ok_lines = int(final_report["Äáº£m báº£o 100%"].sum())
not_ok_lines = total_lines - ok_lines
ok_rate = (ok_lines / total_lines * 100) if total_lines else 0
is_all_ok = total_lines > 0 and not_ok_lines == 0

st.markdown('<div class="step-title">BÆ°á»›c 3: Xem káº¿t luáº­n</div>', unsafe_allow_html=True)
metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("Tá»•ng dÃ²ng", f"{total_lines:,}")
metric2.metric("ÄÃ£ xuáº¥t Ä‘á»§", f"{ok_lines:,}")
metric3.metric("ChÆ°a Ä‘áº£m báº£o", f"{not_ok_lines:,}")
metric4.metric("Tá»· lá»‡ Ä‘áº£m báº£o", f"{ok_rate:.1f}%")

render_result_card(is_all_ok)

error_df = final_report.loc[~final_report["Äáº£m báº£o 100%"], DETAIL_COLUMNS].copy()

if not is_all_ok:
    error_counts = error_df["TÃ¬nh tráº¡ng"].value_counts().rename_axis("TÃ¬nh tráº¡ng").reset_index(name="Sá»‘ dÃ²ng")
    st.dataframe(error_counts, use_container_width=True, hide_index=True, height=150)

    st.dataframe(
        error_df,
        use_container_width=True,
        hide_index=True,
        height=430,
        column_config={
            "Transfer Quantity": st.column_config.NumberColumn("Transfer Quantity", format="%.2f"),
            "Actual Quantity": st.column_config.NumberColumn("Actual Quantity", format="%.2f"),
            "CÃ²n thiáº¿u": st.column_config.NumberColumn("CÃ²n thiáº¿u", format="%.2f"),
            "Gá»£i Ã½ xá»­ lÃ½": st.column_config.TextColumn("Gá»£i Ã½ xá»­ lÃ½", width="large"),
        },
    )

export_bytes = export_excel(final_report, issue_df, mb52_meta)
file_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
st.download_button(
    label="â¬‡ï¸ Táº£i káº¿t quáº£ Excel",
    data=export_bytes,
    file_name=f"StockFlow_KetQua_XuatKho_{file_time}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.caption("StockFlow Checker Â· NgÆ°á»i dÃ¹ng upload phiáº¿u, pháº§n má»m tráº£ lá»i ngay: Ä‘áº£m báº£o 100% hoáº·c thiáº¿u dÃ²ng nÃ o, vÃ¬ sao, xá»­ lÃ½ tháº¿ nÃ o.")
