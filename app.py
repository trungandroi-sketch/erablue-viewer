"""
app.py – Erablue Store Resource Viewer
Reads live data directly from Google Sheets. No local database required.
"""
import io
import datetime
import html as _html
import logging
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px


def search_df(df: pd.DataFrame, query: str, cols: list[str] | None = None) -> pd.DataFrame:
    """Filter DataFrame rows where any of the specified columns contain the query string."""
    if not query:
        return df
    mask = pd.Series([False] * len(df), index=df.index)
    search_cols = cols if cols else list(df.columns[:3])
    for col in search_cols:
        if col in df.columns:
            mask |= df[col].astype(str).str.lower().str.contains(query.lower(), na=False)
    return df[mask]

@st.cache_data(show_spinner=False)
def convert_df_to_excel(df: pd.DataFrame) -> bytes:
    val_io = io.BytesIO()
    with pd.ExcelWriter(val_io, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return val_io.getvalue()

from data_loader import (
    load_sheet, load_erablue,
    refresh, last_refresh_label, SHEET_TABS
)
from html_table_renderer import render_sticky_table, COLUMN_GROUPS
from ai_engine import analyze, BRANDS
import base64

def get_image_base64(path):
    try:
        import os
        if os.path.exists(path):
            with open(path, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
    except Exception:
        pass
    return ""

# ─── Page configuration ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="Erablue Resource Viewer",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Language translation dictionary ──────────────────────────────────────────
T = {
    "vi": {
        "menu_dashboard": "📊 Dashboard",
        "menu_viewer": "📁 Xem Dữ Liệu",
        "menu_reklame": "🎨 Reklame & Branding",
        "menu_ai": "🧠 AI Phân Tích",
        "menu_pending": "📋 Shop Mới",
        "refresh": "🔄 Làm mới dữ liệu",
        "refreshed": "Đã làm mới! Dữ liệu mới nhất từ Google Sheets.",
        "live_status": "✅ Live · GSheets",
        "source_label": "Dữ liệu nguồn: Google Sheets<br>Bấm làm mới để cập nhật dữ liệu mới",
        "offline_status": "⚠️ Ngoại tuyến",
        "db_title": "📊 Dashboard Tổng Quan",
        "db_subtitle": "Báo cáo tài nguyên và độ phủ thương hiệu toàn hệ thống Erablue Electronics",
        "kpi_total": "🏪 Tổng Cửa Hàng",
        "kpi_stores": "cửa hàng trên hệ thống",
        "kpi_area": "🗺️ Khu Vực",
        "kpi_regions": "vùng / {n} tỉnh thành",
        "kpi_best": "🏆 Brand Phủ Nhiều Nhất",
        "kpi_best_desc": "cửa hàng có bàn {brand}",
        "kpi_brands": "📱 Thương Hiệu ICT",
        "kpi_brands_desc": "hãng đang quản lý",
        "coverage_title": "📱 Độ Phủ Thương Hiệu ICT",
        "table_label": "Bàn",
        "wall_label": "Vách",
        "donut_title": "Bàn Demo theo Hãng",
        "dist_title": "🗺️ Phân Bổ Địa Lý",
        "dist_area": "Phân Bổ Cửa Hàng Theo Khu Vực",
        "dist_prov": "Phân Bổ Cửa Hàng Theo Tỉnh / Thành Phố",
        "kpi_opened": "🏪 Đã Khai Trương",
        "kpi_opened_desc": "cửa hàng đã hoạt động",
        "kpi_this_month": "📅 Khai Trương Trong Tháng",
        "kpi_this_month_desc": "trong tháng {month}/{year}",
        "kpi_upcoming": "🚀 Sắp Khai Trương",
        "kpi_upcoming_desc": "dự kiến sắp tới",
        "viewer_title": "📁 Xem Dữ Liệu Chi Tiết",
        "viewer_subtitle": "Dữ liệu đồng bộ trực tiếp từ Google Sheets · Frozen columns · Header phân nhóm theo màu",
        "select_sheet": "📋 Chọn Sheet:",
        "select_sheet_help": "Chọn tên sheet để xem dữ liệu tương ứng từ Google Sheets",
        "search_placeholder": "Nhập để tìm...",
        "search_label": "🔍 Tìm kiếm ID / Tên cửa hàng:",
        "filter_area": "🗺️ Lọc theo Khu vực:",
        "filter_prov": "🏙️ Tỉnh/TP:",
        "partition_label": "🗂️ Phân vùng cột:",
        "partition_help": "Chọn nhóm cột để xem. Mỗi nhóm thể hiện 1 phân khúc tài nguyên.",
        "show_count": "Hiển thị <b>{count}</b> cửa hàng · {cols} cột",
        "no_data": "Không tìm thấy dữ liệu phù hợp với bộ lọc hiện tại.",
        "loading": "Đang tải dữ liệu từ Google Sheets...",
        "loading_sheet": "Đang tải '{sheet}'...",
        "rek_title": "🎨 Reklame & Branding Theo Cửa Hàng",
        "rek_subtitle": "Chi phí bảng hiệu, branding nội thất và thiết bị trưng bày theo từng đối tác nhãn hàng",
        "tab_rek": "🏷️ Reklame Store (Bảng hiệu ngoài)",
        "tab_fix": "🛠️ Fixture Principle (Branding nội thất)",
        "search_shop": "🔍 Tìm cửa hàng:",
        "no_rek": "Không có dữ liệu Reklame.",
        "no_fix": "Không có dữ liệu Fixture Principle.",
        "ai_title": "🧠 AI Phân Tích Dữ Liệu",
        "ai_subtitle": "Hỏi bằng tiếng Việt tự nhiên · AI truy vấn dữ liệu thực từ Google Sheets và trả lời ngay",
        "ai_status": "✅ AI đã kết nối · {count} cửa hàng đang trong cơ sở dữ liệu",
        "ai_no_data": "❌ Không có dữ liệu – AI không thể hoạt động.",
        "ai_presets": "#### 💡 Câu hỏi nhanh:",
        "ai_placeholder": "Ví dụ: Có mấy shop có bàn OPPO? / Shop nào ở Banten? / Tổng hợp độ phủ tất cả hãng? / Phân bổ khu vực?",
        "ai_free_query": "✏️ Hoặc nhập câu hỏi tự do:",
        "ai_analyze_btn": "🔍 Phân Tích Ngay",
        "ai_query_label": "<b>Câu hỏi:</b> {query}",
        "ai_clear_btn": "🗑️ Xóa kết quả",
        "tbl_prov": "Tỉnh / Thành Phố",
        "tbl_count": "Số cửa hàng",
        "export_btn": "📥 Xuất Excel",
        "pending_title": "📋 Danh Sách Shop Mới",
        "pending_subtitle": "Danh sách cửa hàng mới lập, chưa có dữ liệu tài nguyên",
        "pending_count_label": "🏪 Shop mới",
        "pending_surveyed_label": "✅ Đã có dữ liệu",
        "pending_rate_label": "📈 Tỷ lệ phủ dữ liệu",
        "pending_filter_region": "🗺️ Lọc khu vực:",
        "pending_filter_prov": "🏙️ Tỉnh/TP:",
        "pending_search": "🔍 Tìm theo tên shop:",
        "pending_showing": "Hiển thị <b>{n}</b> shop mới",
        "pending_empty": "🎉 Tất cả cửa hàng đều đã được khảo sát & có dữ liệu!",
        "pending_export": "📥 Xuất danh sách Shop mới",
        "pending_col_name": "Tên Cửa Hàng",
        "pending_col_region": "Khu Vực",
        "pending_col_prov": "Tỉnh / TP",
        "pending_col_addr": "Địa Chỉ",
        "pending_col_go": "GO Estimate",
    },
    "en": {
        "menu_dashboard": "📊 Dashboard",
        "menu_viewer": "📁 Data Viewer",
        "menu_reklame": "🎨 Reklame & Branding",
        "menu_ai": "🧠 AI Analyst",
        "menu_pending": "📋 New Stores",
        "refresh": "🔄 Refresh Data",
        "refreshed": "Refreshed! Latest data pulled from Google Sheets.",
        "live_status": "✅ Live · GSheets",
        "source_label": "Source: Google Sheets<br>Click refresh to update data",
        "offline_status": "⚠️ Offline",
        "db_title": "📊 System Overview Dashboard",
        "db_subtitle": "Resource reports and brand coverage across Erablue Electronics",
        "kpi_total": "🏪 Total Stores",
        "kpi_stores": "stores in system",
        "kpi_area": "🗺️ Regions",
        "kpi_regions": "regions / {n} provinces",
        "kpi_best": "🏆 Top Covered Brand",
        "kpi_best_desc": "stores with {brand} table",
        "kpi_brands": "📱 ICT Brands",
        "kpi_brands_desc": "managed brands",
        "coverage_title": "📱 ICT Brand Coverage",
        "table_label": "Table",
        "wall_label": "Wall",
        "donut_title": "Demo Table by Brand",
        "dist_title": "🗺️ Geographical Distribution",
        "dist_area": "Stores by Region",
        "dist_prov": "Stores by Province / City",
        "kpi_opened": "🏪 Opened Stores",
        "kpi_opened_desc": "active stores",
        "kpi_this_month": "📅 Opened This Month",
        "kpi_this_month_desc": "in {month}/{year}",
        "kpi_upcoming": "🚀 Coming Soon",
        "kpi_upcoming_desc": "scheduled openings",
        "viewer_title": "📁 Detailed Data Viewer",
        "viewer_subtitle": "Synchronized directly from Google Sheets · Frozen columns · Grouped colored headers",
        "select_sheet": "📋 Select Sheet:",
        "select_sheet_help": "Select sheet to view data from Google Sheets",
        "search_placeholder": "Type to search...",
        "search_label": "🔍 Search ID / Shop Name:",
        "filter_area": "🗺️ Filter by Region:",
        "filter_prov": "🏙️ Province/City:",
        "partition_label": "🗂️ Column Partition:",
        "partition_help": "Select a subset of columns to display.",
        "show_count": "Showing <b>{count}</b> stores · {cols} columns",
        "no_data": "No matching stores found for current filters.",
        "loading": "Loading data from Google Sheets...",
        "loading_sheet": "Loading '{sheet}'...",
        "rek_title": "🎨 Reklame & Branding by Store",
        "rek_subtitle": "Outdoor signage costs, interior branding, and display fixtures by brand",
        "tab_rek": "🏷️ Reklame Store (Outdoor Signage)",
        "tab_fix": "🛠️ Fixture Principle (Interior Branding)",
        "search_shop": "🔍 Search store:",
        "no_rek": "No Reklame data available.",
        "no_fix": "No Fixture Principle data available.",
        "ai_title": "🧠 AI Data Analyst",
        "ai_subtitle": "Ask in natural language · AI queries Google Sheets data and responds instantly",
        "ai_status": "✅ AI Connected · {count} stores in database",
        "ai_no_data": "❌ No data loaded – AI cannot operate.",
        "ai_presets": "#### 💡 Quick Questions:",
        "ai_placeholder": "E.g. How many stores have OPPO tables? / Which stores are in Banten? / Summary of all brand coverage?",
        "ai_free_query": "✏️ Ask anything:",
        "ai_analyze_btn": "🔍 Analyze Now",
        "ai_query_label": "<b>Question:</b> {query}",
        "ai_clear_btn": "🗑️ Clear Result",
        "tbl_prov": "Province / City",
        "tbl_count": "Stores",
        "export_btn": "📥 Export to Excel",
        "pending_title": "📋 New Stores List",
        "pending_subtitle": "Newly imported stores with no resource data yet",
        "pending_count_label": "🏪 New Stores",
        "pending_surveyed_label": "✅ Has Resource Data",
        "pending_rate_label": "📈 Data Coverage Rate",
        "pending_filter_region": "🗺️ Filter Region:",
        "pending_filter_prov": "🏙️ Province/City:",
        "pending_search": "🔍 Search by name:",
        "pending_showing": "Showing <b>{n}</b> new stores",
        "pending_empty": "🎉 All stores have been surveyed & have data!",
        "pending_export": "📥 Export New Stores to Excel",
        "pending_col_name": "Store Name",
        "pending_col_region": "Region",
        "pending_col_prov": "Province / City",
        "pending_col_addr": "Address",
        "pending_col_go": "GO Estimate",
    }
}

# ─── Global CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,300;0,14..32,400;0,14..32,500;0,14..32,600;0,14..32,700;0,14..32,800;1,14..32,400&display=swap');

*, *::before, *::after {
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
    text-rendering: optimizeLegibility;
    box-sizing: border-box;
}

html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
    background-color: #f3f4f6 !important;
    color: #1e293b !important;
    font-size: 15px; /* Increased from 14px */
    line-height: 1.65;
    letter-spacing: -0.015em;
}

/* Hide default Streamlit chrome, but keep header visible for the sidebar toggle button */
#MainMenu, footer { display: none !important; visibility: hidden !important; }
header { background: transparent !important; }
.block-container { padding: 2rem 2.5rem 3rem 2.5rem !important; max-width: 100% !important; }

/* Hide top-right developer toolbar elements (Share, Edit, Star, GitHub link) for a clean UI */
[data-testid="stHeader"] button:not([data-testid="stSidebarCollapseButton"]):not([aria-label="Expand sidebar"]) { display: none !important; }
[data-testid="stHeader"] a { display: none !important; }

/* Hide Streamlit Community Cloud and Hugging Face floating developer/hosting badges at the bottom right */
[data-testid="stViewerBadge"],
.viewerBadge,
[data-testid="stAppShareButton"],
div[class*="styles_viewerBadge"],
div[class*="viewerBadge"],
#ConnectionStatus {
    display: none !important;
}

/* ── FIX: Sidebar always expanded – hide ALL toggle buttons, force translateX(0) ── */
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
[aria-label="Collapse sidebar"],
[aria-label="Expand sidebar"] {
    display: none !important;
    visibility: hidden !important;
    pointer-events: none !important;
}

/* Force sidebar to always remain visible even when browser cache has it collapsed */
section[data-testid="stSidebar"] {
    transform: translateX(0) !important;
    visibility: visible !important;
    display: flex !important;
    min-width: 15rem !important;
    max-width: 22rem !important;
    width: 21.5rem !important;
    transition: none !important;
}
section[data-testid="stSidebar"] > div:first-child {
    width: 21.5rem !important;
}

/* Sidebar styling with rich royal violet gradient matching Creative Tim */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #8b5cf6 0%, #5b21b6 100%) !important;
    border-right: 1px solid rgba(255, 255, 255, 0.05) !important;
}
[data-testid="stSidebar"] * { color: #f1f5f9 !important; }
[data-testid="stSidebar"] button {
    background-color: rgba(255, 255, 255, 0.15) !important;
    color: #ffffff !important;
    border: 1px solid rgba(255, 255, 255, 0.25) !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    transition: all 0.25s ease !important;
}
[data-testid="stSidebar"] button:hover {
    background-color: rgba(255, 255, 255, 0.25) !important;
    border-color: rgba(255, 255, 255, 0.4) !important;
    color: #ffffff !important;
}
[data-testid="stSidebar"] button p {
    color: #ffffff !important;
}
[data-testid="stSidebar"] .stRadio label { color: #cbd5e1 !important; font-size: 15px; }
[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p { color: #94a3b8 !important; font-size: 12px; }

/* Page animations (smooth slide & fade transition) */
@keyframes slideFade {
    from { opacity: 0; transform: translateY(12px); }
    to   { opacity: 1; transform: translateY(0); }
}
.page-wrap { 
    animation: slideFade 0.45s cubic-bezier(0.16, 1, 0.3, 1) both; 
    will-change: transform, opacity;
}

/* Micro-animations and transitions for premium feel */
button, [data-testid="stDownloadButton"] button, .stButton button, [data-testid="stBaseButton-secondary"] {
    transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 14.5px !important;
}
button:hover, [data-testid="stDownloadButton"] button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(15, 39, 68, 0.15) !important;
}
button:active, [data-testid="stDownloadButton"] button:active {
    transform: scale(0.98) !important;
}

/* KPI Cards: Premium Glassmorphism */
.kpi-card {
    background: rgba(255, 255, 255, 0.85);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border: 1px solid rgba(255, 255, 255, 0.5);
    border-radius: 16px;
    padding: 22px 26px;
    box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04), 0 2px 6px rgba(15, 23, 42, 0.02);
    border-left: 5px solid #2563eb;
    transition: all 0.35s cubic-bezier(0.16, 1, 0.3, 1);
}
.kpi-card:hover { 
    transform: translateY(-3px); 
    box-shadow: 0 10px 30px rgba(15, 23, 42, 0.08), 0 4px 12px rgba(15, 23, 42, 0.03);
    border-left-width: 7px;
}
.kpi-label { 
    font-size: 12px; 
    font-weight: 700; 
    letter-spacing: .8px;
    text-transform: uppercase; 
    color: #64748b; 
    margin-bottom: 5px; 
}
.kpi-value { 
    font-size: 2.8rem; /* Increased from 2.4rem */
    font-weight: 850; 
    color: #0f172a; 
    line-height: 1.15; 
}
.kpi-sub { 
    font-size: 13px; 
    color: #64748b; 
    margin-top: 6px; 
}

/* Brand coverage table: premium rows */
.brand-row {
    display: flex; 
    align-items: center; 
    gap: 14px;
    padding: 12px 18px; 
    border-radius: 12px;
    background: rgba(255, 255, 255, 0.8);
    backdrop-filter: blur(8px);
    -webkit-backdrop-filter: blur(8px);
    margin-bottom: 8px;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.4);
    border-left: 5px solid var(--brand-color);
    transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
}
.brand-row:hover {
    transform: translateX(3px);
    box-shadow: 0 4px 15px rgba(15, 23, 42, 0.06);
    background: #ffffff;
}
.brand-bar-bg {
    flex: 1; 
    height: 10px; 
    background: #e2e8f0;
    border-radius: 6px; 
    overflow: hidden;
}
.brand-bar-fill {
    height: 100%; 
    border-radius: 6px;
    background: var(--brand-color);
    transition: width 0.8s cubic-bezier(0.16, 1, 0.3, 1);
}

/* Section headers */
.section-title {
    font-size: 20px; /* Increased from 18px */
    font-weight: 800; 
    color: #0f172a;
    margin: 2rem 0 1rem; 
    padding-bottom: .5rem;
    border-bottom: 2px solid #e2e8f0;
}

/* Search & filter bar */
.filter-bar {
    background: rgba(255, 255, 255, 0.9);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255, 255, 255, 0.5);
    border-radius: 12px; 
    padding: 14px 18px;
    box-shadow: 0 4px 15px rgba(15, 23, 42, 0.03); 
    margin-bottom: 16px;
}

/* Page header: Premium clean white card with violet accent bar matching Creative Tim */
.page-header {
    background: #ffffff !important;
    border-radius: 16px; 
    padding: 24px 32px; 
    margin-bottom: 28px;
    color: #0f172a !important;
    box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04), 0 2px 6px rgba(15, 23, 42, 0.02);
    border: 1px solid #e2e8f0;
    border-top: 4px solid #8b5cf6; /* Royal violet accent bar matching the sidebar */
}
.page-header h1 { 
    font-size: 27px;
    font-weight: 850; 
    margin: 0; 
    color: #0f172a !important; 
    letter-spacing: -0.02em;
}
.page-header p { 
    font-size: 14px;
    color: #64748b !important; 
    margin: 6px 0 0; 
    line-height: 1.5;
}

/* Partition pills */
.stSelectbox label { 
    font-weight: 700; 
    font-size: 14px; 
    color: #0f172a; 
}

/* Data badge */
.data-badge {
    display: inline-flex; 
    align-items: center; 
    gap: 6px;
    background: #dcfce7; 
    color: #166534; 
    padding: 5px 14px;
    border-radius: 20px; 
    font-size: 12px; 
    font-weight: 700;
    box-shadow: 0 1px 3px rgba(22, 101, 52, 0.1);
}
.data-badge-warn {
    background: #fef9c3; 
    color: #854d0e;
    box-shadow: 0 1px 3px rgba(133, 77, 14, 0.1);
}

/* Style Streamlit radio buttons as premium navigation menus in the sidebar */
[data-testid="stRadio"] div[role="radiogroup"] {
    gap: 10px !important;
    background: transparent !important;
}
[data-testid="stRadio"] div[role="radiogroup"] > label > div:first-child {
    display: none !important;
}
[data-testid="stRadio"] div[role="radiogroup"] > label {
    background-color: rgba(255, 255, 255, 0.03) !important;
    border: 1px solid rgba(255, 255, 255, 0.06) !important;
    padding: 12px 16px !important;
    border-radius: 10px !important;
    color: #e2e8f0 !important;
    cursor: pointer !important;
    transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    min-height: 46px !important;
    margin: 0 !important;
    flex: 1 !important;
    width: 100% !important;
}
[data-testid="stRadio"] div[role="radiogroup"] > label:hover {
    background-color: rgba(255, 255, 255, 0.1) !important;
    color: #ffffff !important;
    border-color: rgba(255, 255, 255, 0.15) !important;
}
[data-testid="stRadio"] div[role="radiogroup"] > label:has(input:checked) {
    background: rgba(255, 255, 255, 0.16) !important;
    border-color: rgba(255, 255, 255, 0.35) !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06) !important;
    color: #ffffff !important;
}
[data-testid="stRadio"] div[role="radiogroup"] > label:has(input:checked) [data-testid="stMarkdownContainer"] p {
    color: #ffffff !important;
    font-weight: 700 !important;
}
[data-testid="stRadio"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"] p {
    font-size: 13px !important;
    margin: 0 !important;
    padding: 0 !important;
}
</style>
""", unsafe_allow_html=True)


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_base64 = get_image_base64("logo.png")
    if logo_base64:
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="width: 100px; height: 100px; border-radius: 50%; box-shadow: 0 4px 15px rgba(0,0,0,0.25); border: 2px solid rgba(255,255,255,0.25); margin-bottom: 12px;">'
    else:
        logo_html = '<div style="font-size:36px;">⚡</div>'
        
    st.markdown(f"""
    <div style="text-align:center; padding: 16px 0 12px;">
        {logo_html}
        <div style="font-size:18px;font-weight:800;color:#e2e8f0;">Erablue</div>
        <div style="font-size:11px;color:#cbd5e1;letter-spacing:.5px;text-transform:uppercase;margin-top:4px;">
            Resource Viewer
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Language Toggle VNI / English
    lang_choice = st.radio(
        "Language",
        ["Tiếng Việt 🇻🇳", "English 🇬🇧"],
        horizontal=True,
        label_visibility="collapsed"
    )
    lang = "vi" if "Việt" in lang_choice else "en"

    st.markdown("<hr style='border-color:rgba(255,255,255,.1);margin:0 0 16px;'>", unsafe_allow_html=True)

    menu_opts = [
        T[lang]["menu_dashboard"],
        T[lang]["menu_viewer"],
        T[lang]["menu_reklame"],
        T[lang]["menu_ai"],
        T[lang]["menu_pending"],
    ]
    menu_sel = st.radio(
        "**MENU**",
        menu_opts,
        label_visibility="visible",
    )
    menu = ["dashboard", "viewer", "reklame", "ai", "pending"][menu_opts.index(menu_sel)]

    st.markdown("<hr style='border-color:rgba(255,255,255,.1);margin:12px 0;'>", unsafe_allow_html=True)

    # Refresh button
    if st.button(T[lang]["refresh"], use_container_width=True):
        refresh()
        st.success(T[lang]["refreshed"])
        st.rerun()

    # Load status
    try:
        ts = last_refresh_label()
        st.markdown(
            f'<div class="data-badge">{T[lang]["live_status"]}</div>'
            f'<div style="font-size:10px;color:#64748b;margin-top:4px;">{ts}</div>',
            unsafe_allow_html=True,
        )
    except Exception as e:
        logging.exception(f"Sidebar status display error: {e}")
        st.markdown(f'<div class="data-badge data-badge-warn">{T[lang]["offline_status"]}</div>',
                    unsafe_allow_html=True)

    st.markdown(f"""
    <div style="text-align:center; margin-top: 35px; padding-bottom: 20px;">
        <div style="font-size:10px;color:#94a3b8;">
            {T[lang]["source_label"]}
        </div>
    </div>
    """, unsafe_allow_html=True)


# ─── Load data ────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def get_erablue():
    return load_erablue()

with st.spinner(T[lang]["loading"]):
    try:
        df_main = get_erablue()
        data_ok = True
    except Exception as e:
        st.error(f"❌ Error: {e}")
        data_ok = False
        df_main = pd.DataFrame()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 – DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
if menu == "dashboard":
    st.markdown('<div class="page-wrap">', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="page-header">
        <h1>{T[lang]["db_title"]}</h1>
        <p>{T[lang]["db_subtitle"]}</p>
    </div>
    """, unsafe_allow_html=True)

    if not data_ok or df_main.empty:
        st.warning("No data available.")
        st.stop()

    total = len(df_main)
    area_col = next((c for c in df_main.columns if "Khu" in c and "vực" in c.lower()), None)
    prov_col = next((c for c in df_main.columns if "Rút gọn" in c), None) or next((c for c in df_main.columns if "Tỉnh" in c), None)
    n_areas = df_main[area_col].nunique() if area_col else 0
    n_provinces = df_main[prov_col].nunique() if prov_col else 0

    # 1. Identify resource columns start to find unsurveyed stores (stores with 0.0 resource data)
    RESOURCE_ANCHORS = ['bàn demo', 'vách tivi', 'tv đảo', 'máy lạnh principle', 'tủ lạnh principle', 'máy giặt principle']
    resource_col_start = None
    for i, col in enumerate(df_main.columns):
        if any(kw in col.lower() for kw in RESOURCE_ANCHORS):
            resource_col_start = i
            break
    if resource_col_start is None:
        resource_col_start = 11

    resource_section = df_main.iloc[:, resource_col_start:]
    resource_num = resource_section.apply(pd.to_numeric, errors='coerce').fillna(0)
    row_resource_sum = resource_num.sum(axis=1)
    is_surveyed = row_resource_sum > 0
    unsurveyed = int((~is_surveyed).sum())

    # 2. Dynamic brand detection from 'Bàn Demo' columns
    BRAND_COLORS = {
        "samsung": "#1428a0", "apple": "#3a3a3c", "oppo": "#1a7c40",
        "xiaomi": "#e05c00", "vivo": "#3251ff", "realme": "#c49000",
        "infinix": "#52606d", "huawei": "#10b981", "itel": "#ef4444",
        "tecno": "#0284c7"
    }

    brand_stats = []
    for col in df_main.columns:
        if "bàn demo" in col.lower() or "ban demo" in col.lower():
            # Clean and split suffix to extract brand name
            b_name = col
            for sfx in ["bàn demo", "ban demo"]:
                idx = b_name.lower().find(sfx)
                if idx != -1:
                    b_name = b_name[:idx].strip()
                    break
            
            display_name = b_name
            if display_name.lower().startswith("apple"):
                display_name = "Apple"
                
            brand_key = display_name.lower()
            
            # Find the corresponding Wall column
            wall_col = None
            for w_col in df_main.columns:
                w_lower = w_col.lower()
                if b_name.lower() in w_lower and ("tường" in w_lower or "tuong" in w_lower or "wall" in w_lower):
                    wall_col = w_col
                    break
            
            tc_num = pd.to_numeric(df_main[col], errors='coerce').fillna(0)
            tc = int((tc_num > 0).sum())
            tc_no = int(((tc_num <= 0) & is_surveyed).sum())
            
            wc = 0
            wc_no = 0
            if wall_col:
                wc_num = pd.to_numeric(df_main[wall_col], errors='coerce').fillna(0)
                wc = int((wc_num > 0).sum())
                wc_no = int(((wc_num <= 0) & is_surveyed).sum())
                
            color = BRAND_COLORS.get(brand_key, "#64748b")
            brand_stats.append({
                "name": display_name,
                "color": color,
                "table": tc,
                "table_no": tc_no,
                "wall": wc,
                "wall_no": wc_no,
                "wall_col": wall_col,
                "pct": round(tc / total * 100, 1) if total else 0,
            })

    best_brand = max(brand_stats, key=lambda x: x["table"]) if brand_stats else {"table": 0, "name": "N/A", "color": "#64748b"}

    # ── Calculate opening statuses ─────────────────────────────────────────
    go_col = next((c for c in df_main.columns if "ước tính" in c.lower() and "go" in c.lower()), None)
    
    n_opened = total
    n_this_month = 0
    n_upcoming = 0
    
    if go_col:
        # Preprocess date string: replace "xx"/Draft days with "01" to make it parseable for month/year comparisons
        raw_series = df_main[go_col].astype(str).str.strip()
        clean_series = raw_series.str.replace(r'^[xX]{2}([/\-])', r'01\1', regex=True)
        
        go_series = pd.to_datetime(clean_series, dayfirst=True, errors="coerce")
        today = datetime.date.today()
        
        # A valid full date must match DD/MM/YYYY or YYYY-MM-DD and NOT start with "xx" (case-insensitive)
        date_only_str = raw_series.str.split().str[0]
        is_full_date = (
            (date_only_str.str.match(r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{4}$', na=False) |
             date_only_str.str.match(r'^\d{4}[/\-]\d{1,2}[/\-]\d{1,2}$', na=False)) &
            (~raw_series.str.lower().str.startswith("xx"))
        )
        
        valid_mask = go_series.notna() & (go_series.dt.year > 1900)
        
        # Sắp khai trương: GO date in the future months (after current month/year)
        upcoming_mask = valid_mask & (
            (go_series.dt.year > today.year) |
            ((go_series.dt.year == today.year) & (go_series.dt.month > today.month))
        )
        n_upcoming = int(upcoming_mask.sum())
        
        # Khai trương trong tháng: GO date in current month/year
        this_month_mask = valid_mask & (go_series.dt.year == today.year) & (go_series.dt.month == today.month)
        n_this_month = int(this_month_mask.sum())
        
        # Đã khai trương: GO date is in the past (<= today) AND must be a full valid date (no "xx" or partial month)
        # OR has a valid Ngày Setup date in the past when GO date is blank (excluding xx and blank setup dates)
        opened_setup = pd.Series([False] * len(df_main), index=df_main.index)
        setup_col = next((c for c in df_main.columns if "setup" in c.lower() and "ngày" in c.lower()), None)
        if setup_col:
            raw_setup = df_main[setup_col].astype(str).str.strip()
            setup_series = pd.to_datetime(raw_setup, dayfirst=True, errors="coerce")
            setup_only_str = raw_setup.str.split().str[0]
            is_full_setup = (
                (setup_only_str.str.match(r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{4}$', na=False) |
                 setup_only_str.str.match(r'^\d{4}[/\-]\d{1,2}[/\-]\d{1,2}$', na=False)) &
                (~raw_setup.str.lower().str.startswith("xx"))
            )
            valid_setup = setup_series.notna() & (setup_series.dt.year > 1900)
            opened_setup = (df_main[go_col].isna() | (df_main[go_col].astype(str).str.strip() == "")) & \
                           valid_setup & (setup_series <= pd.Timestamp(today)) & is_full_setup
        
        opened_mask = (valid_mask & (go_series <= pd.Timestamp(today)) & is_full_date) | opened_setup
        n_opened = int(opened_mask.sum())

    # ── KPI row ────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    kpis = [
        (c1, T[lang]["kpi_total"], total, T[lang]["kpi_stores"], "#2563eb"),
        (c2, T[lang]["kpi_area"], n_areas, T[lang]["kpi_regions"].format(n=n_provinces), "#0891b2"),
        (c3, T[lang]["kpi_best"], best_brand["table"],
         T[lang]["kpi_best_desc"].format(brand=best_brand["name"]), best_brand["color"]),
        (c4, T[lang]["kpi_brands"], len(brand_stats), T[lang]["kpi_brands_desc"], "#7c3aed"),
    ]
    for col, label, val, sub, color in kpis:
        with col:
            st.markdown(
                f'<div class="kpi-card" style="border-left-color:{color};">'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{val}</div>'
                f'<div class="kpi-sub">{sub}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    # ── KPI row 2 (Operation Status) ───────────────────────────────────────
    st.markdown('<div style="margin-top:12px;margin-bottom:8px;"></div>', unsafe_allow_html=True)
    c1_2, c2_2, c3_2 = st.columns(3)
    
    today = datetime.date.today()
    upcoming_val = n_upcoming if n_upcoming > 0 else (
        "Đang cập nhật" if lang == "vi" else "Updating"
    )
    kpis_row2 = [
        (c1_2, T[lang]["kpi_opened"], n_opened, T[lang]["kpi_opened_desc"], "#16a34a"),
        (c2_2, T[lang]["kpi_this_month"], n_this_month, T[lang]["kpi_this_month_desc"].format(month=today.month, year=today.year), "#ea580c"),
        (c3_2, T[lang]["kpi_upcoming"], upcoming_val, T[lang]["kpi_upcoming_desc"], "#2563eb"),
    ]
    for col, label, val, sub, color in kpis_row2:
        val_str = str(val)
        # If the value is a text string (not a simple number), use a smaller font size to fit well
        font_size = "1.5rem" if any(c.isalpha() for c in val_str) else "2rem"
        with col:
            st.markdown(
                f'<div class="kpi-card" style="border-left-color:{color}; padding: 16px 20px;">'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value" style="font-size:{font_size};color:{color};">{val}</div>'
                f'<div class="kpi-sub">{sub}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown(f'<div class="section-title">{T[lang]["coverage_title"]}</div>',
                unsafe_allow_html=True)

    # ── Brand coverage table ───────────────────────────────────────────────
    col_left, col_right = st.columns([2, 1])

    with col_left:
        for b in sorted(brand_stats, key=lambda x: -x["table"]):
            pct = b["pct"]
            
            # Wall display HTML
            if b["wall_col"]:
                wall_html = (
                    f'<div style="min-width: 140px; text-align: left;">'
                    f'  🧱 <b>{T[lang]["wall_label"]}:</b> {b["wall"]} có | {total - b["wall"]} chưa có<br>'
                    f'  <span style="font-size:10px;color:#64748b;">(gồm <b style="color:#d97706;">{b["wall_no"]}</b> đã k.sát, <b style="color:#64748b;">{unsurveyed}</b> shop mới)</span>'
                    f'</div>'
                )
            else:
                wall_html = f'<div style="min-width: 140px; text-align: left; color:#94a3b8;">🧱 <b>{T[lang]["wall_label"]}:</b> Không quản lý</div>'
                
            st.markdown(
                f'<div class="brand-row" style="--brand-color:{b["color"]}; flex-direction: column; align-items: stretch; padding: 10px 14px; height: auto; margin-bottom: 6px;">'
                f'  <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">'
                f'    <div style="min-width:90px;font-weight:700;font-size:12.5px;">{b["name"]}</div>'
                f'    <div class="brand-bar-bg" style="flex-grow: 1; margin: 0 12px;"><div class="brand-bar-fill" style="width:{pct}%;"></div></div>'
                f'    <div style="min-width:38px;text-align:right;font-weight:700;font-size:12.5px;color:{b["color"]};">{pct}%</div>'
                f'  </div>'
                f'  <div style="display: flex; font-size: 11px; color: #475569; justify-content: space-between; border-top: 1px solid #f1f5f9; padding-top: 4px; line-height: 1.4;">'
                f'    <div style="min-width: 140px; text-align: left;">'
                f'      📱 <b>{T[lang]["table_label"]}:</b> {b["table"]} có | {total - b["table"]} chưa có<br>'
                f'      <span style="font-size:10px;color:#64748b;">(gồm <b style="color:#d97706;">{b["table_no"]}</b> đã k.sát, <b style="color:#64748b;">{unsurveyed}</b> shop mới)</span>'
                f'    </div>'
                f'    {wall_html}'
                f'  </div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    with col_right:
        # Plotly donut for brand coverage
        fig = go.Figure(go.Pie(
            labels=[b["name"] for b in brand_stats],
            values=[b["table"] for b in brand_stats],
            hole=0.55,
            marker_colors=[b["color"] for b in brand_stats],
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>%{value} stores<br>%{percent}<extra></extra>",
        ))
        fig.update_layout(
            showlegend=True,
            legend=dict(orientation="v", font=dict(size=11)),
            margin=dict(l=10, r=10, t=30, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=280,
            title=dict(text=T[lang]["donut_title"], font=dict(size=13, color="#0f2744")),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # ── CE Brand Coverage (position-based, avoids cross-category name collisions) ─
    BRAND_COLORS = {
        "Sony":      "#1c1c1c", "Samsung":   "#1428a0", "Polytron":  "#dc2626",
        "Sharp":     "#374151", "Toshiba":   "#9f1239", "TCL":       "#d97706",
        "Panasonic": "#003087", "Daikin":    "#005bac", "LG":        "#a50034",
        "Midea":     "#0073b7", "Gree":      "#16a34a", "Aqua":      "#0ea5e9",
        "Electrolux":"#6d28d9",
    }

    # Anchor keywords identify the FIRST column of each CE category group
    CE_ANCHOR_DEFS = [
        ("vach_tivi", ["vách tivi", "vach tivi"]),
        ("tv_dao",    ["tv đảo", "tv dao"]),
        ("may_lanh",  ["máy lạnh principle", "may lanh principle"]),
        ("tu_lanh",   ["tủ lạnh principle", "tu lanh principle"]),
        ("may_giat",  ["máy giặt principle", "may giat principle"]),
        # Terminator: stops may_giat range before SDA/poster columns that contain brand names
        ("_end_ce",   ["đầu tiên của dòng", "first tv of line", "kệ sda",
                       "first shelf", "sda -", "sda–", "poster", "tổng poster"]),
    ]
    # Brands expected in each group (for counting)
    CE_CAT_BRANDS = {
        "vach_tivi": ["Sony", "Samsung", "Polytron", "Sharp", "Toshiba", "TCL"],
        "tv_dao":    ["Samsung", "Sharp", "Sony", "Polytron", "Toshiba", "TCL"],
        "may_lanh":  ["Panasonic", "Daikin", "LG", "Samsung", "Polytron",
                      "Sharp", "Midea", "Gree", "Aqua", "TCL", "Electrolux"],
        "tu_lanh":   ["Midea", "TCL", "Aqua", "Polytron", "Sharp", "Toshiba"],
        "may_giat":  ["Midea", "Aqua", "Polytron", "Sharp", "Toshiba"],
    }

    def _ce_brand_counts_positional(df: pd.DataFrame) -> dict:
        """
        Use anchor column positions + column ranges to count CE brand stores.
        Avoids cross-category collisions (e.g. 'Samsung' appears in TV đảo AND Máy lạnh).
        Returns: {cat_key: {brand_name: count}}
        """
        cols = list(df.columns)
        n    = len(cols)

        # 1. Find the index of the first column in each CE category
        anchor_idx = {}
        for cat, kws in CE_ANCHOR_DEFS:
            for i, col in enumerate(cols):
                if any(kw in col.lower() for kw in kws):
                    anchor_idx[cat] = i
                    break

        if not anchor_idx:
            return {}

        # 2. Determine column ranges [start, end) per category
        sorted_cats = sorted(anchor_idx.items(), key=lambda x: x[1])
        ranges = {}
        for k, (cat, start) in enumerate(sorted_cats):
            end = sorted_cats[k + 1][1] if k + 1 < len(sorted_cats) else n
            ranges[cat] = (start, end)

        # 3. For each category × brand, count shops with value > 0
        #    Use FIRST-match-then-break strategy:
        #    - CE data columns appear immediately after their anchor
        #    - Later columns (poster, SDA…) may share brand names → skip them
        #      by stopping at the first positive column match per brand.
        results = {}
        for cat, (start, end) in ranges.items():
            if cat.startswith("_"):  # terminator sentinel, skip
                continue
            cat_cols = cols[start:end]
            brand_counts = {}
            for brand in CE_CAT_BRANDS.get(cat, []):
                brand_l = brand.lower()
                count   = 0
                for col in cat_cols:
                    cl = col.lower()
                    if "vị trí" in cl or "vi tri" in cl:   # skip text/location cols
                        continue
                    if brand_l in cl:
                        num = pd.to_numeric(df[col], errors="coerce").fillna(0)
                        count = int((num > 0).sum())
                        break  # ← take FIRST matching col only; prevents false
                               #   positives from poster/SDA cols further in the range
                brand_counts[brand] = count
            results[cat] = brand_counts
        return results

    ce_counts = _ce_brand_counts_positional(df_main)

    # ── Tivi = Vách Tivi + TV Đảo (union by brand, take max) ─────────────────
    tivi_brands_set = ["Sony", "Samsung", "Polytron", "Sharp", "Toshiba", "TCL"]
    vach = ce_counts.get("vach_tivi", {})
    dao  = ce_counts.get("tv_dao",    {})
    tivi_brands = [
        {"name": b, "color": BRAND_COLORS.get(b, "#64748b"),
         "count": max(vach.get(b, 0), dao.get(b, 0))}
        for b in tivi_brands_set
    ]

    # ── Máy Lạnh, Tủ Lạnh, Máy Giặt ─────────────────────────────────────────
    def _cat_to_list(cat_key):
        data = ce_counts.get(cat_key, {})
        return [
            {"name": b, "color": BRAND_COLORS.get(b, "#64748b"), "count": data.get(b, 0)}
            for b in CE_CAT_BRANDS.get(cat_key, [])
        ]

    may_lanh_brands = _cat_to_list("may_lanh")
    tu_lanh_brands  = _cat_to_list("tu_lanh")
    may_giat_brands = _cat_to_list("may_giat")

    # ── Render section ────────────────────────────────────────────────────────
    ce_title   = "📺 Độ Phủ Thương Hiệu CE"  if lang == "vi" else "📺 CE Brand Coverage"
    tivi_lbl   = "📺 Tivi Hãng"              if lang == "vi" else "📺 TV Brands"
    ac_lbl     = "❄️ Máy Lạnh Hãng"         if lang == "vi" else "❄️ AC Brands"
    fridge_lbl = "🧊 Tủ Lạnh Hãng"          if lang == "vi" else "🧊 Fridge Brands"
    wm_lbl     = "🫧 Máy Giặt Hãng"          if lang == "vi" else "🫧 Washer Brands"

    st.markdown(f'<div class="section-title">{ce_title}</div>', unsafe_allow_html=True)

    ce_col1, ce_col2, ce_col3, ce_col4 = st.columns(4)

    def _render_ce_col(col_ctx, brand_list, title):
        with col_ctx:
            st.markdown(
                f'<div style="font-weight:700;font-size:13px;color:#0f2744;margin-bottom:10px;'
                f'padding:8px 12px;background:#f1f5f9;border-radius:8px;">{title}</div>',
                unsafe_allow_html=True,
            )
            has_data = any(b["count"] > 0 for b in brand_list)
            if not has_data:
                st.markdown(
                    '<div style="color:#94a3b8;font-size:12px;padding:6px 12px;">'
                    + ("Chưa có dữ liệu" if lang == "vi" else "No data") + "</div>",
                    unsafe_allow_html=True,
                )
                return
            for b in sorted(brand_list, key=lambda x: -x["count"]):
                pct = round(b["count"] / total * 100, 1) if total else 0
                st.markdown(
                    f'<div class="brand-row" style="--brand-color:{b["color"]};'
                    f'padding:7px 10px;margin-bottom:4px;">'
                    f'  <div style="min-width:72px;font-weight:700;font-size:11.5px;">{b["name"]}</div>'
                    f'  <div class="brand-bar-bg"><div class="brand-bar-fill" style="width:{pct}%;"></div></div>'
                    f'  <div style="min-width:38px;text-align:right;font-weight:700;'
                    f'font-size:11.5px;color:{b["color"]}">{pct}%</div>'
                    f'  <div style="min-width:52px;font-size:10.5px;color:#64748b;">'
                    f'{b["count"]} shops</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    _render_ce_col(ce_col1, tivi_brands,      tivi_lbl)
    _render_ce_col(ce_col2, may_lanh_brands,  ac_lbl)
    _render_ce_col(ce_col3, tu_lanh_brands,   fridge_lbl)
    _render_ce_col(ce_col4, may_giat_brands,  wm_lbl)

    # ── Geographical Distribution row ──────────────────────────────────────
    st.markdown(f'<div class="section-title">{T[lang]["dist_title"]}</div>', unsafe_allow_html=True)
    dcol1, dcol2 = st.columns(2)

    with dcol1:
        st.markdown(f'<div style="font-weight:600;font-size:13px;color:#0f2744;margin-bottom:8px;">{T[lang]["dist_area"]}</div>', unsafe_allow_html=True)
        if area_col:
            by_area = (
                df_main.groupby(area_col).size()
                .reset_index(name="count")
                .sort_values("count", ascending=True)
            )
            fig2 = px.bar(
                by_area, x="count", y=area_col,
                orientation="h",
                color="count",
                color_continuous_scale=["#bfdbfe", "#1d4ed8"],
                labels={"count": T[lang]["kpi_stores"], area_col: ""},
                text="count",
            )
            fig2.update_traces(textposition="outside")
            fig2.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
                coloraxis_showscale=False,
                margin=dict(l=10, r=40, t=10, b=10),
                height=300,
                xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
                yaxis=dict(tickfont=dict(size=12)),
            )
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    with dcol2:
        st.markdown(f'<div style="font-weight:600;font-size:13px;color:#0f2744;margin-bottom:8px;">{T[lang]["dist_prov"]}</div>', unsafe_allow_html=True)
        if prov_col:
            by_prov = (
                df_main.groupby(prov_col).size()
                .reset_index(name="count")
                .sort_values("count", ascending=False)
            )
            rows_html = []
            for _, r in by_prov.iterrows():
                prov_safe = _html.escape(str(r[prov_col]))
                count_safe = _html.escape(str(r["count"]))
                rows_html.append(
                    f'<tr style="border-bottom:1px solid #f1f5f9; color:#334155;">'
                    f'  <td style="padding:10px 0; font-weight:500;">{prov_safe}</td>'
                    f'  <td style="padding:10px 0; text-align:right; font-weight:600; color:#ea580c;">{count_safe}</td>'
                    f'</tr>'
                )
            rows_str = "\n".join(rows_html)
            
            table_html = f"""
            <div style="max-height:280px; overflow-y:auto; border:1px solid #e2e8f0; border-radius:10px; background:#ffffff; padding:0 12px; margin-top:2px;">
                <table style="width:100%; border-collapse:collapse; font-size:12.5px;">
                    <thead>
                        <tr style="border-bottom:2px solid #e2e8f0; position:sticky; top:0; background:#ffffff; z-index:1; color:#0f2744; font-weight:600;">
                            <th style="padding:10px 0; text-align:left; font-size:13px;">{T[lang]["tbl_prov"]}</th>
                            <th style="padding:10px 0; text-align:right; font-size:13px;">{T[lang]["tbl_count"]}</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows_str}
                    </tbody>
                </table>
            </div>
            """
            st.markdown(table_html, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 – DATA VIEWER
# ═══════════════════════════════════════════════════════════════════════════════
elif menu == "viewer":
    st.markdown('<div class="page-wrap">', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="page-header">
        <h1>{T[lang]["viewer_title"]}</h1>
        <p>{T[lang]["viewer_subtitle"]}</p>
    </div>
    """, unsafe_allow_html=True)

    # Sheet selector (tabs)
    sheet_choice = st.selectbox(
        T[lang]["select_sheet"],
        SHEET_TABS,
        help=T[lang]["select_sheet_help"],
    )

    with st.spinner(T[lang]["loading_sheet"].format(sheet=sheet_choice)):
        try:
            if sheet_choice == "Erablue Existing":
                df_view = df_main.copy()
            else:
                df_view = load_sheet(sheet_choice)
        except Exception as e:
            st.error(f"Lỗi tải dữ liệu: {e}")
            st.stop()

    # ── Filter bar ─────────────────────────────────────────────────────────
    with st.container():
        fcol1, fcol2, fcol3 = st.columns([2, 2, 1])
        with fcol1:
            search = st.text_input(T[lang]["search_label"], placeholder=T[lang]["search_placeholder"])
        with fcol2:
            if "Khu vực" in df_view.columns and sheet_choice == "Erablue Existing":
                areas = ["Tất cả" if lang == "vi" else "All"] + sorted(df_view["Khu vực"].dropna().unique().tolist())
                area_filter = st.selectbox(T[lang]["filter_area"], areas)
            elif "Area" in df_view.columns and sheet_choice == "Erablue Existing":
                areas = ["Tất cả" if lang == "vi" else "All"] + sorted(df_view["Area"].dropna().unique().tolist())
                area_filter = st.selectbox(T[lang]["filter_area"], areas)
            else:
                area_filter = "Tất cả" if lang == "vi" else "All"
        with fcol3:
            if "Tỉnh/Thành phố (Rút gọn)" in df_view.columns and sheet_choice == "Erablue Existing":
                provs = ["Tất cả" if lang == "vi" else "All"] + sorted(df_view["Tỉnh/Thành phố (Rút gọn)"].dropna().unique().tolist())
                prov_filter = st.selectbox(T[lang]["filter_prov"], provs)
            elif "Province/City Simplified" in df_view.columns and sheet_choice == "Erablue Existing":
                provs = ["Tất cả" if lang == "vi" else "All"] + sorted(df_view["Province/City Simplified"].dropna().unique().tolist())
                prov_filter = st.selectbox(T[lang]["filter_prov"], provs)
            else:
                prov_filter = "Tất cả" if lang == "vi" else "All"

    # Apply filters
    filtered = df_view.copy()
    if search:
        filtered = search_df(filtered, search, ["ID Cửa hàng", "Tên Cửa hàng", "ID Store", "Shop Name", "Store ID", "ID store", "ID STORE", "Tên cửa hàng", "Shop name"])
        
    if area_filter not in ["Tất cả", "All"]:
        if "Khu vực" in filtered.columns:
            filtered = filtered[filtered["Khu vực"] == area_filter]
        elif "Area" in filtered.columns:
            filtered = filtered[filtered["Area"] == area_filter]
            
    if prov_filter not in ["Tất cả", "All"]:
        if "Tỉnh/Thành phố (Rút gọn)" in filtered.columns:
            filtered = filtered[filtered["Tỉnh/Thành phố (Rút gọn)"] == prov_filter]
        elif "Province/City Simplified" in filtered.columns:
            filtered = filtered[filtered["Province/City Simplified"] == prov_filter]

    # Partitions block
    # Partitions block - column lists defined once, labels vary by language
    _PARTITION_DEFS = [
        ("all", "🌐 Tất cả cột", "🌐 All Columns", None),
        ("capacity", "📊 Tổng công suất", "📊 Total Capacity", [
            "Còn lại Bàn", "Tường",
            "Tài nguyên Layout Bàn", "Tường.1",
            "Tài nguyên Thực tế Bàn", "Tường.2",
        ]),
        ("ict", "📱 ICT – Bàn & Vách (tất cả hãng)", "📱 ICT – Table & Wall (All brands)", [
            "Samsung Bàn Demo", "Tủ Tường Thương hiệu",
            "Apple 1.2m Bàn Demo", "Tủ Tường Thương hiệu.1",
            "OPPO Bàn Demo", "Tủ Tường Thương hiệu.2",
            "Xiaomi Bàn Demo", "Tủ Tường Thương hiệu.3",
            "Vivo Bàn Demo", "Tủ Tường Thương hiệu.4",
            "Realme Bàn Demo", "Tủ Tường Thương hiệu.5",
            "Đa thương hiệu (Huawei, Realme, Infinix) Demo ĐA THƯƠNG HIỆU ( Infinix )",
            "GHI CHÚ",
        ]),
        ("acc_laptop", "💻 Khu Vực Phụ Kiện & Laptop", "💻 Accessories & Laptop Zone", [
            "Tủ Điện thoại Điện thoại", "Máy tính bảng",
            "LAPTOP PRINCIPLE HP", "INTEL", "ACER", "ASUS", "LENOVO",
            "Laptop Bàn Laptop", "Tủ Laptop",
            "Phụ kiện Tường", "Fixture (Sàn)", "CCTV + Loa",
        ]),
        ("erablue", "⚡ Tài Nguyên Erablue Electronics", "⚡ Erablue Electronics Resources", [
            "Tài nguyên cho Erablue Electronics TV Treo tường",
            "TV Bàn", "Tủ đông", "Nền Tủ lạnh", "Tủ lạnh Tường",
            "Nền Máy giặt", "Nền Máy sấy", "Nền Máy rửa chén",
            "KỆ MÁY GIẶT", "Máy giặt Tường", "Máy lạnh Tường",
            "RIG", "Máy nước nóng Tường",
        ]),
        ("tv_wall", "📺 TV Treo Tường (Vị Trí Ưu Tiên)", "📺 Brand TV – Wall (Priority Loc)", [
            "TV Treo tường (Vị trí ưu tiên) Vị trí Sony", "Sony (m)",
            "Vị trí Samsung", "Samsung (m)",
            "Vị trí Polytron", "Polytron",
            "Vị trí Sharp", "Sharp",
            "Vị trí Toshiba", "Toshiba",
            "Vị trí TCL", "TCL",
        ]),
        ("tv_island", "📺 TV Đảo (Island TV)", "📺 Island TV", [
            "TV đảo principle (Nguyên tắc) (/Kệ) Samsung",
            "Sharp .1", "Sony", "Polytron .1", "Toshiba .1", "TCL .1",
        ]),
        ("ac_fridge", "❌️ Máy Lạnh & Tủ Lạnh Hãng", "❌️ AC & Fridge by Brand", [
            "Máy lạnh principle (SL) Panasonic",
            "Daikin", "LG", "Samsung",
            "Polytron .2", "Sharp .2", "Midea", "Gree",
            "Aqua", "TCL .2", "Electrolux",
            "Tủ lạnh principle (/Kệ) Midea",
            "TCL .3", "Aqua .1", "Polytron .3", "Sharp .3", "Toshiba .2",
        ]),
        ("wm_sda", "🦺 WM Đảo & SDA", "🦺 WM Island & SDA", [
            "MÁY GIAT PRINCIPLE Midea",
            "TCL .4", "Aqua .2", "Polytron .4", "Sharp .4", "Toshiba .3",
            "TV đầu tiên của dòng", "Kệ SDA đầu tiên",
            "SDA MIYAKO (TƯỜNG)", "MIYAKO (ENDCAP)",
            "PHILIPS (TƯỜNG)", "PHILIPS (ENDCAP)",
            "ELECTROLUX (TƯỜNG)", "ELECTROLUX (ENDCAP)",
            "MIDEA (TƯỜNG)", "MIDEA (ENDCAP)", "Kệ", "Tường.3",
        ]),
        ("poster", "🖼️ Poster & Mặt Tiền & Diện Tích", "🖼️ Poster & Facade & Area", [
            "TỔNG POSTER TƯỜNG SỬ DỤNG", "CÒN LẠI",
            "POSTER TƯỜNG Thuê theo Thương hiệu",
            "Samsung.1", "Aqua.1", "Polytron.1", "LG.1", "Elux", "Sharp.1",
            "Logo Erablue", "Logo Erafone",
            "Mặt tiền Chính (C)", "Khác (R)", "Khác (L)",
            "Diện tích (m2) Kho Điện máy",
            "WC + Phòng Nhân viên", "Kho + Server", "Bãi đậu xe",
            "Showroom", "Tổng diện tích", "Đất trống",
        ]),
    ]
    PARTITIONS = {}
    for _, vi_lbl, en_lbl, cols in _PARTITION_DEFS:
        label = en_lbl if lang == "en" else vi_lbl
        PARTITIONS[label] = cols

    if sheet_choice == "Erablue Existing":
        pcol1, pcol2 = st.columns([2, 1])
        with pcol1:
            part_key = st.selectbox(
                T[lang]["partition_label"],
                list(PARTITIONS.keys()),
                help=T[lang]["partition_help"],
            )
            part_cols = PARTITIONS[part_key]
        with pcol2:
            limit_choice = st.selectbox(
                "⚡ Tốc độ tải (Số dòng):" if lang == "vi" else "⚡ Load speed (Rows):",
                [50, 100, 200, "Hiển thị tất cả" if lang == "vi" else "Show all"],
                index=0,
                help="Giới hạn số dòng hiển thị của bảng lớn để tối ưu tốc độ load và hiệu năng trình duyệt." if lang == "vi" else "Limits rows to maximize render speed and browser performance."
            )
    else:
        part_key = ""
        part_cols = None
        limit_choice = st.selectbox(
            "⚡ Tốc độ tải (Số dòng):" if lang == "vi" else "⚡ Load speed (Rows):",
            [50, 100, 200, "Hiển thị tất cả" if lang == "vi" else "Show all"],
            index=0,
            help="Giới hạn số dòng hiển thị của bảng lớn để tối ưu tốc độ load và hiệu năng trình duyệt." if lang == "vi" else "Limits rows to maximize render speed and browser performance."
        )

    # Build display dataframe
    CORE_ID_COLS = ["ID Cửa hàng", "Tên Cửa hàng", "Tỉnh/Thành phố (Rút gọn)", "Kích thước Cửa hàng", "Khu vực"]
    if part_cols is not None:
        id_present = [c for c in CORE_ID_COLS if c in filtered.columns]
        extra = [c for c in part_cols if c in filtered.columns and c not in id_present]
        display_df = filtered[id_present + extra]
    else:
        display_df = filtered

    total_rows = len(display_df)
    
    # Slice dataframe if speed limit is active
    if limit_choice not in ["Hiển thị tất cả", "Show all"]:
        display_df_limited = display_df.head(int(limit_choice))
        limited_active = True
    else:
        display_df_limited = display_df
        limited_active = False

    # Display row count and Export button
    dcol_left, dcol_right = st.columns([3, 1])
    with dcol_left:
        limit_text = f" (đã tối ưu hiển thị {limit_choice} dòng)" if limited_active and total_rows > int(limit_choice) else ""
        limit_text_en = f" (displaying first {limit_choice} rows)" if limited_active and total_rows > int(limit_choice) else ""
        st.markdown(
            f'<div style="padding-top:8px;font-size:13px;color:#64748b;">'
            f'{T[lang]["show_count"].format(count=total_rows, cols=len(display_df.columns))}{limit_text if lang == "vi" else limit_text_en}</div>',
            unsafe_allow_html=True,
        )
    with dcol_right:
        if not display_df.empty:
            xlsx_data = convert_df_to_excel(display_df)
            st.download_button(
                label=T[lang]["export_btn"],
                data=xlsx_data,
                file_name=f"{sheet_choice}_export_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="viewer_export_btn",
                use_container_width=True
            )

    # Render the sticky HTML table
    if not display_df_limited.empty:
        html_out = render_sticky_table(display_df_limited, max_height=820)
        components.html(html_out, height=875, scrolling=False)
    else:
        st.info(T[lang]["no_data"])

    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 – REKLAME & BRANDING
# ═══════════════════════════════════════════════════════════════════════════════
elif menu == "reklame":
    st.markdown('<div class="page-wrap">', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="page-header">
        <h1>{T[lang]["rek_title"]}</h1>
        <p>{T[lang]["rek_subtitle"]}</p>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner(T[lang]["loading_sheet"].format(sheet="Fixture principle")):
        try:
            df_fix = load_sheet("Fixture principle")
        except Exception as e:
            st.error(str(e))
            df_fix = pd.DataFrame()
    if not df_fix.empty:
        scol1, scol2 = st.columns([3, 1])
        with scol1:
            search_fix = st.text_input(T[lang]["search_shop"], key="fix_search")
        if search_fix:
            df_fix = search_df(df_fix, search_fix)
        with scol2:
            st.markdown('<div style="height:28px;"></div>', unsafe_allow_html=True)
            xlsx_fix = convert_df_to_excel(df_fix)
            st.download_button(
                label=T[lang]["export_btn"],
                data=xlsx_fix,
                file_name=f"Fixture_principle_export_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="fix_export_btn",
                use_container_width=True
            )
        html_fix = render_sticky_table(df_fix, max_height=700)
        components.html(html_fix, height=750, scrolling=False)
    else:
        st.warning(T[lang]["no_fix"])

    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 – AI ANALYST
# ═══════════════════════════════════════════════════════════════════════════════
elif menu == "ai":
    st.markdown('<div class="page-wrap">', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="page-header">
        <h1>{T[lang]["ai_title"]}</h1>
        <p>{T[lang]["ai_subtitle"]}</p>
    </div>
    """, unsafe_allow_html=True)

    # Status
    if data_ok:
        st.success(T[lang]["ai_status"].format(count=len(df_main)))
    else:
        st.error(T[lang]["ai_no_data"])
        st.stop()

    # Quick-action buttons
    st.markdown(T[lang]["ai_presets"])
    qcols = st.columns(4)
    presets = [
        ("📱 Bàn OPPO", "Có mấy shop có bàn OPPO?"),
        ("📱 Vách Samsung", "Có mấy shop có vách Samsung?"),
        ("🍎 Bàn Apple", "Có mấy shop có bàn Apple?"),
        ("📱 Bàn Xiaomi", "Có mấy shop có bàn Xiaomi?"),
    ]
    qcols2 = st.columns(4)
    presets2 = [
        ("📱 Bàn Vivo", "Có mấy shop có bàn Vivo?"),
        ("📱 Bàn Realme", "Có mấy shop có bàn Realme?"),
        ("🗺️ Phân bổ vùng", "Phân bổ cửa hàng theo khu vực?"),
        ("📊 Tất cả hãng", "Tổng hợp độ phủ tất cả hãng ICT?"),
    ]

    # Initialize history
    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []

    # Handle preset quick-actions
    run_preset = None
    for col, (label, query) in zip(qcols, presets):
        with col:
            if st.button(label, use_container_width=True):
                run_preset = query

    for col, (label, query) in zip(qcols2, presets2):
        with col:
            if st.button(label, use_container_width=True):
                run_preset = query

    st.markdown("---")

    # Render previous chat history
    chat_container = st.container()
    with chat_container:
        for msg in st.session_state["chat_history"]:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

    # Input handling
    user_query = st.chat_input(T[lang]["ai_placeholder"])
    if run_preset:
        user_query = run_preset

    if user_query:
        # Add user query to chat history
        st.session_state["chat_history"].append({"role": "user", "content": user_query})
        with chat_container:
            with st.chat_message("user"):
                st.write(user_query)
            
            with st.chat_message("assistant"):
                with st.spinner("AI is analyzing live data..." if lang == "vi" else "AI is analyzing live data..."):
                    try:
                        response_text = analyze(user_query, df_main, lang=lang)
                    except Exception as e:
                        logging.error(f"AI query analysis error: {str(e)}")
                        if lang == "vi":
                            response_text = "❌ **Đã xảy ra lỗi:** Em không thể biên dịch hoặc phân tích câu hỏi này. Anh vui lòng diễn đạt lại câu hỏi rõ ràng hơn nhé!"
                        else:
                            response_text = "❌ **Error:** I could not compile or analyze this query. Please rephrase your question more clearly!"
                    st.write(response_text)
                    
        # Add assistant response to history
        st.session_state["chat_history"].append({"role": "assistant", "content": response_text})
        st.rerun()

    # Clear chat button
    if st.session_state["chat_history"]:
        st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
        if st.button(T[lang]["ai_clear_btn"]):
            st.session_state["chat_history"] = []
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 – STORES PENDING SURVEY
# ═══════════════════════════════════════════════════════════════════════════════
if menu == "pending":
    st.markdown('<div class="page-wrap">', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="page-header">
        <h1>{T[lang]["pending_title"]}</h1>
        <p>{T[lang]["pending_subtitle"]}</p>
    </div>
    """, unsafe_allow_html=True)

    if not data_ok or df_main.empty:
        st.warning("No data available.")
        st.stop()

    # ── Detect resource columns (CE + ICT anchor columns) ─────────────────────
    RESOURCE_ANCHORS = [
        'bàn demo', 'vách tivi', 'tv đảo', 'máy lạnh principle',
        'tủ lạnh principle', 'máy giặt principle',
        'wall tv', 'island tv', 'wall air',
    ]
    resource_col_start = None
    for i, col in enumerate(df_main.columns):
        if any(kw in col.lower() for kw in RESOURCE_ANCHORS):
            resource_col_start = i
            break

    if resource_col_start is None:
        # Fallback: skip first 8 assumed-basic columns
        resource_col_start = min(8, len(df_main.columns) - 1)

    # Numeric sum of resource section per row
    resource_section = df_main.iloc[:, resource_col_start:]
    resource_num     = resource_section.apply(pd.to_numeric, errors='coerce').fillna(0)
    row_resource_sum = resource_num.sum(axis=1)

    pending_mask  = row_resource_sum == 0
    df_pending    = df_main[pending_mask].copy()
    df_surveyed   = df_main[~pending_mask]

    total_stores   = len(df_main)
    n_pending      = len(df_pending)
    n_surveyed     = len(df_surveyed)
    coverage_rate  = round(n_surveyed / total_stores * 100, 1) if total_stores else 0

    # ── KPI row ───────────────────────────────────────────────────────────────
    k1, k2, k3 = st.columns(3)
    with k1:
        st.markdown(
            f'<div class="kpi-card" style="border-left-color:#f59e0b;">'
            f'<div class="kpi-label">{T[lang]["pending_count_label"]}</div>'
            f'<div class="kpi-value" style="color:#d97706;">{n_pending:,}</div>'
            f'<div class="kpi-sub">/ {total_stores:,} tổng số cửa hàng</div></div>',
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f'<div class="kpi-card" style="border-left-color:#22c55e;">'
            f'<div class="kpi-label">{T[lang]["pending_surveyed_label"]}</div>'
            f'<div class="kpi-value" style="color:#16a34a;">{n_surveyed:,}</div>'
            f'<div class="kpi-sub">cửa hàng đã khảo sát xong</div></div>',
            unsafe_allow_html=True,
        )
    with k3:
        bar_color = "#22c55e" if coverage_rate >= 70 else ("#f59e0b" if coverage_rate >= 40 else "#ef4444")
        st.markdown(
            f'<div class="kpi-card" style="border-left-color:{bar_color};">'
            f'<div class="kpi-label">{T[lang]["pending_rate_label"]}</div>'
            f'<div class="kpi-value" style="color:{bar_color};">{coverage_rate}%</div>'
            f'<div class="kpi-sub">cửa hàng đã có dữ liệu tài nguyên</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # Progress bar
    st.markdown(
        f'<div style="background:#e2e8f0;border-radius:999px;height:10px;overflow:hidden;margin-bottom:20px;">'
        f'<div style="width:{coverage_rate}%;height:100%;background:linear-gradient(90deg,#22c55e,#16a34a);'
        f'border-radius:999px;transition:width .6s ease;"></div></div>',
        unsafe_allow_html=True,
    )

    if n_pending == 0:
        st.success(T[lang]["pending_empty"])
        st.stop()

    # ── Detect basic info columns ─────────────────────────────────────────────
    basic_cols = list(df_main.columns[:resource_col_start])

    # Try to identify specific useful columns
    name_col   = next((c for c in basic_cols if any(k in c.lower() for k in
                       ['tên', 'name', 'store'])), basic_cols[0] if basic_cols else None)
    addr_col   = next((c for c in basic_cols if any(k in c.lower() for k in
                       ['địa chỉ', 'address', 'alamat'])), None)
    region_col = next((c for c in basic_cols if 'khu' in c.lower() and 'vực' in c.lower()), None)
    prov_col   = next((c for c in basic_cols if 'rút gọn' in c.lower()), None) or \
                 next((c for c in basic_cols if 'tỉnh' in c.lower()), None)
    go_col     = next((c for c in basic_cols if 'go' in c.lower() and
                       any(k in c.lower() for k in ['estimate', 'date', 'khai'])), None) or \
                 next((c for c in basic_cols if 'estimate' in c.lower()), None)

    display_cols = [c for c in [name_col, region_col, prov_col, addr_col, go_col] if c]
    if not display_cols:
        display_cols = basic_cols[:6]

    # ── Filters ───────────────────────────────────────────────────────────────
    fa, fb, fc = st.columns([1.5, 1.5, 2])

    with fa:
        if region_col and df_pending[region_col].notna().any():
            all_regions = ["Tất cả"] + sorted(df_pending[region_col].dropna().unique().tolist())
            sel_region = st.selectbox(T[lang]["pending_filter_region"], all_regions)
        else:
            sel_region = "Tất cả"

    with fb:
        if prov_col and df_pending[prov_col].notna().any():
            source_df = df_pending if sel_region == "Tất cả" else df_pending[
                df_pending[region_col] == sel_region] if region_col else df_pending
            all_provs = ["Tất cả"] + sorted(source_df[prov_col].dropna().unique().tolist())
            sel_prov = st.selectbox(T[lang]["pending_filter_prov"], all_provs)
        else:
            sel_prov = "Tất cả"

    with fc:
        search_q = st.text_input(T[lang]["pending_search"], placeholder="...")

    # ── Apply filters ─────────────────────────────────────────────────────────
    df_view = df_pending.copy()
    if sel_region != "Tất cả" and region_col:
        df_view = df_view[df_view[region_col] == sel_region]
    if sel_prov != "Tất cả" and prov_col:
        df_view = df_view[df_view[prov_col] == sel_prov]
    if search_q and name_col:
        df_view = df_view[df_view[name_col].astype(str).str.lower()
                          .str.contains(search_q.lower(), na=False)]

    # ── Summary bar ───────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="filter-bar" style="display:flex;justify-content:space-between;align-items:center;">'
        f'<span style="font-size:13px;color:#334155;">{T[lang]["pending_showing"].format(n=len(df_view))}</span>',
        unsafe_allow_html=True,
    )

    # ── Export button ─────────────────────────────────────────────────────────
    if not df_view.empty:
        export_df = df_view[display_cols].rename(columns={
            name_col:   T[lang]["pending_col_name"]   if name_col   else "Name",
            region_col: T[lang]["pending_col_region"] if region_col else "Region",
            prov_col:   T[lang]["pending_col_prov"]   if prov_col   else "Province",
            addr_col:   T[lang]["pending_col_addr"]   if addr_col   else "Address",
            go_col:     T[lang]["pending_col_go"]     if go_col     else "GO",
        }) if all(c in df_view.columns for c in display_cols if c) else df_view.iloc[:, :6]

        excel_bytes = convert_df_to_excel(export_df)
        st.download_button(
            label=T[lang]["pending_export"],
            data=excel_bytes,
            file_name=f"pending_survey_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Table ─────────────────────────────────────────────────────────────────
    if df_view.empty:
        st.info(T[lang]["no_data"] if "no_data" in T[lang] else "No stores found.")
    else:
        # Style table rows
        st.markdown("""
        <style>
        .pending-table { width:100%; border-collapse:collapse; font-size:13px;
            font-family:'Inter',-apple-system,sans-serif; }
        .pending-table th { background:#0f2744; color:#e2e8f0; padding:10px 14px;
            text-align:left; font-weight:600; font-size:11px;
            text-transform:uppercase; letter-spacing:.5px; position:sticky;top:0; }
        .pending-table td { padding:9px 14px; border-bottom:1px solid #f1f5f9;
            color:#334155; vertical-align:top; }
        .pending-table tr:hover td { background:#f8fafc; }
        .pending-table tr:nth-child(even) td { background:#fafafa; }
        .badge-region { display:inline-block; padding:2px 8px; border-radius:12px;
            font-size:10.5px; font-weight:600; background:#dbeafe; color:#1e40af; }
        .badge-prov   { display:inline-block; padding:2px 8px; border-radius:12px;
            font-size:10.5px; font-weight:600; background:#dcfce7; color:#15803d; }
        .pending-wrap { max-height:60vh; overflow-y:auto; border-radius:10px;
            box-shadow:0 2px 12px rgba(0,0,0,.08); }
        </style>
        """, unsafe_allow_html=True)

        rows_html = ""
        for _, row in df_view.iterrows():
            name_val   = str(row[name_col])   if name_col   and name_col   in row.index else "—"
            region_val = str(row[region_col]) if region_col and region_col in row.index else "—"
            prov_val   = str(row[prov_col])   if prov_col   and prov_col   in row.index else "—"
            addr_val   = str(row[addr_col])   if addr_col   and addr_col   in row.index else "—"
            go_val     = str(row[go_col])     if go_col     and go_col     in row.index else "—"

            rows_html += (
                f"<tr>"
                f"<td><b>{_html.escape(name_val)}</b></td>"
                f"<td><span class='badge-region'>{_html.escape(region_val)}</span></td>"
                f"<td><span class='badge-prov'>{_html.escape(prov_val)}</span></td>"
                f"<td style='color:#64748b;font-size:12px;'>{_html.escape(addr_val)}</td>"
                f"<td style='color:#94a3b8;font-size:12px;'>{_html.escape(go_val)}</td>"
                f"</tr>"
            )

        th = T[lang]
        table_html = f"""
        <div class='pending-wrap'>
        <table class='pending-table'>
        <thead><tr>
            <th>{th['pending_col_name']}</th>
            <th>{th['pending_col_region']}</th>
            <th>{th['pending_col_prov']}</th>
            <th>{th['pending_col_addr']}</th>
            <th>{th['pending_col_go']}</th>
        </tr></thead>
        <tbody>{rows_html}</tbody>
        </table></div>
        """
        st.markdown(table_html, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)
