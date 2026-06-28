"""
app.py – Erablue Store Resource Viewer
Reads live data directly from Google Sheets. No local database required.
"""
import io
import datetime
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

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
        "refresh": "🔄 Làm mới dữ liệu",
        "refreshed": "Đã làm mới! Dữ liệu mới nhất từ Google Sheets.",
        "live_status": "✅ Live · GSheets",
        "source_label": "Dữ liệu nguồn: Google Sheets<br>Tự động làm mới mỗi 5 phút",
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
    },
    "en": {
        "menu_dashboard": "📊 Dashboard",
        "menu_viewer": "📁 Data Viewer",
        "menu_reklame": "🎨 Reklame & Branding",
        "menu_ai": "🧠 AI Analyst",
        "refresh": "🔄 Refresh Data",
        "refreshed": "Refreshed! Latest data pulled from Google Sheets.",
        "live_status": "✅ Live · GSheets",
        "source_label": "Source: Google Sheets<br>Auto-refresh every 5 minutes",
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
    }
}

# ─── Global CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif !important;
    background-color: #f8fafc !important;
    color: #0f172a !important;
}

/* Hide default Streamlit chrome, but keep header visible for the sidebar toggle button */
#MainMenu, footer { display: none !important; visibility: hidden !important; }
header { background: transparent !important; }
.block-container { padding: 1.5rem 2rem 2rem 2rem !important; max-width: 100% !important; }

/* Hide top-right developer toolbar elements (Share, Edit, Star, GitHub link) for a clean UI */
[data-testid="stHeader"] button { display: none !important; }
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

/* Custom premium styling for sidebar toggle button when collapsed */
[data-testid="stHeader"] [data-testid="collapsedControl"],
[data-testid="stHeader"] [data-testid="collapsedControl"] button {
    display: inline-flex !important;
    background-color: #0f2744 !important;
    border-radius: 0 8px 8px 0 !important;
    color: #ffffff !important;
    box-shadow: 2px 0 8px rgba(15, 39, 68, 0.2) !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    transition: background-color 0.2s !important;
}
[data-testid="stHeader"] [data-testid="collapsedControl"]:hover,
[data-testid="stHeader"] [data-testid="collapsedControl"] button:hover {
    background-color: #1d4ed8 !important;
}
[data-testid="stHeader"] [data-testid="collapsedControl"] svg {
    color: #ffffff !important;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f2744 0%, #1a3a6e 100%) !important;
    border-right: none !important;
}
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
[data-testid="stSidebar"] .stRadio label { color: #cbd5e1 !important; font-size: 14px; }
[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p { color: #94a3b8 !important; font-size: 11px; }

/* Page animations */
@keyframes fadeUp {
    from { opacity: 0; transform: translateY(10px); }
    to   { opacity: 1; transform: translateY(0); }
}
.page-wrap { animation: fadeUp .35s ease both; }

/* KPI Cards */
.kpi-card {
    background: #ffffff;
    border-radius: 14px;
    padding: 20px 24px;
    box-shadow: 0 2px 12px rgba(30,58,95,.08);
    border-left: 4px solid #2563eb;
    transition: transform .2s, box-shadow .2s;
}
.kpi-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(30,58,95,.13); }
.kpi-label { font-size: 11px; font-weight: 600; letter-spacing: .6px;
    text-transform: uppercase; color: #64748b; margin-bottom: 4px; }
.kpi-value { font-size: 2.4rem; font-weight: 800; color: #0f2744; line-height: 1.1; }
.kpi-sub   { font-size: 12px; color: #94a3b8; margin-top: 4px; }

/* Brand coverage table */
.brand-row {
    display: flex; align-items: center; gap: 12px;
    padding: 10px 16px; border-radius: 8px;
    background: #fff; margin-bottom: 6px;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
    border-left: 4px solid var(--brand-color);
}
.brand-bar-bg {
    flex: 1; height: 8px; background: #e2e8f0;
    border-radius: 4px; overflow: hidden;
}
.brand-bar-fill {
    height: 100%; border-radius: 4px;
    background: var(--brand-color);
    transition: width .6s ease;
}

/* Section headers */
.section-title {
    font-size: 18px; font-weight: 700; color: #0f2744;
    margin: 1.5rem 0 .8rem; padding-bottom: .4rem;
    border-bottom: 2px solid #e2e8f0;
}

/* Search & filter bar */
.filter-bar {
    background: #fff; border-radius: 10px; padding: 12px 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,.06); margin-bottom: 12px;
}

/* Page header */
.page-header {
    background: linear-gradient(135deg, #0f2744 0%, #1d4ed8 100%);
    border-radius: 14px; padding: 24px 32px; margin-bottom: 24px;
    color: white;
}
.page-header h1 { font-size: 26px; font-weight: 800; margin: 0; color: white; }
.page-header p  { font-size: 13px; color: #93c5fd; margin: 6px 0 0; }

/* Partition pills */
.stSelectbox label { font-weight: 600; font-size: 13px; color: #1e3a5f; }

/* Data badge */
.data-badge {
    display: inline-flex; align-items: center; gap: 6px;
    background: #dcfce7; color: #15803d; padding: 4px 12px;
    border-radius: 20px; font-size: 11px; font-weight: 600;
}
.data-badge-warn {
    background: #fef9c3; color: #a16207;
}
</style>
""", unsafe_allow_html=True)


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding: 16px 0 12px;">
        <div style="font-size:36px;">⚡</div>
        <div style="font-size:18px;font-weight:800;color:#e2e8f0;">Erablue</div>
        <div style="font-size:11px;color:#64748b;letter-spacing:.5px;text-transform:uppercase;">
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
        T[lang]["menu_ai"]
    ]
    menu_sel = st.radio(
        "**MENU**",
        menu_opts,
        label_visibility="visible",
    )
    menu = ["dashboard", "viewer", "reklame", "ai"][menu_opts.index(menu_sel)]

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
    except Exception:
        st.markdown(f'<div class="data-badge data-badge-warn">{T[lang]["offline_status"]}</div>',
                    unsafe_allow_html=True)

    st.markdown(f"""
    <div style="position:absolute;bottom:16px;left:0;right:0;text-align:center;">
        <div style="font-size:10px;color:#475569;">
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

    # Compute brand stats
    brand_stats = []
    from ai_engine import BRANDS, _count_positive
    for b in BRANDS[:6]:  # ICT brands only
        tc, _ = _count_positive(df_main, b["table"])
        wc = 0
        if b["wall"]:
            wc, _ = _count_positive(df_main, b["wall"])
        brand_stats.append({
            "name": b["name"], "color": b["color"],
            "table": tc, "wall": wc,
            "pct": round(tc / total * 100, 1) if total else 0,
        })

    best_brand = max(brand_stats, key=lambda x: x["table"])

    # ── Calculate opening statuses ─────────────────────────────────────────
    go_col = next((c for c in df_main.columns if "ước tính" in c.lower() and "go" in c.lower()), None)
    
    n_opened = total
    n_this_month = 0
    n_upcoming = 0
    
    if go_col:
        go_series = pd.to_datetime(df_main[go_col], errors="coerce")
        today = datetime.date.today()
        
        valid_mask = go_series.notna() & (go_series.dt.year > 1900)
        
        # Sắp khai trương: GO date in the future
        upcoming_mask = valid_mask & (go_series.dt.date > today)
        n_upcoming = int(upcoming_mask.sum())
        
        # Khai trương trong tháng: GO date in current month/year
        this_month_mask = valid_mask & (go_series.dt.year == today.year) & (go_series.dt.month == today.month)
        n_this_month = int(this_month_mask.sum())
        
        # Đã khai trương: GO date is in the past (<= today)
        opened_mask = valid_mask & (go_series.dt.date <= today)
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
    kpis_row2 = [
        (c1_2, T[lang]["kpi_opened"], n_opened, T[lang]["kpi_opened_desc"], "#16a34a"),
        (c2_2, T[lang]["kpi_this_month"], n_this_month, T[lang]["kpi_this_month_desc"].format(month=today.month, year=today.year), "#ea580c"),
        (c3_2, T[lang]["kpi_upcoming"], n_upcoming, T[lang]["kpi_upcoming_desc"], "#2563eb"),
    ]
    for col, label, val, sub, color in kpis_row2:
        with col:
            st.markdown(
                f'<div class="kpi-card" style="border-left-color:{color}; padding: 16px 20px;">'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value" style="font-size:2rem;color:{color};">{val}</div>'
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
            st.markdown(
                f'<div class="brand-row" style="--brand-color:{b["color"]};">'
                f'  <div style="min-width:90px;font-weight:700;font-size:13px;">{b["name"]}</div>'
                f'  <div class="brand-bar-bg"><div class="brand-bar-fill" style="width:{pct}%;"></div></div>'
                f'  <div style="min-width:38px;text-align:right;font-weight:700;font-size:13px;color:{b["color"]};">{pct}%</div>'
                f'  <div style="min-width:80px;font-size:11px;color:#64748b;">{T[lang]["table_label"]}: {b["table"]}</div>'
                f'  <div style="min-width:80px;font-size:11px;color:#64748b;">{T[lang]["wall_label"]}: {b["wall"]}</div>'
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
                rows_html.append(
                    f'<tr style="border-bottom:1px solid #f1f5f9; color:#334155;">'
                    f'  <td style="padding:10px 0; font-weight:500;">{r[prov_col]}</td>'
                    f'  <td style="padding:10px 0; text-align:right; font-weight:600; color:#ea580c;">{r["count"]}</td>'
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
        mask = pd.Series([False] * len(filtered), index=filtered.index)
        for col in ["ID Cửa hàng", "Tên Cửa hàng", "ID Store", "Shop Name", "Store ID", "ID store", "ID STORE", "Tên cửa hàng", "Shop name"]:
            if col in filtered.columns:
                mask |= filtered[col].astype(str).str.lower().str.contains(search.lower(), na=False)
        filtered = filtered[mask]
        
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
    if lang == "vi":
        PARTITIONS = {
            "🌐 Tất cả cột": None,
            "📊 Tổng công suất": [
                "Còn lại Bàn", "Tường",
                "Tài nguyên Layout Bàn", "Tường.1",
                "Tài nguyên Thực tế Bàn", "Tường.2",
            ],
            "📱 ICT – Bàn & Vách (tất cả hãng)": [
                "Samsung Bàn Demo", "Tủ Tường Thương hiệu",
                "Apple 1.2m Bàn Demo", "Tủ Tường Thương hiệu.1",
                "OPPO Bàn Demo", "Tủ Tường Thương hiệu.2",
                "Xiaomi Bàn Demo", "Tủ Tường Thương hiệu.3",
                "Vivo Bàn Demo", "Tủ Tường Thương hiệu.4",
                "Realme Bàn Demo", "Tủ Tường Thương hiệu.5",
                "Đa thương hiệu (Huawei, Realme, Infinix) Demo ĐA THƯƠNG HIỆU ( Infinix )",
                "GHI CHÚ",
            ],
            "⚡ Tài Nguyên Erablue Electronics": [
                "Tài nguyên cho Erablue Electronics TV Treo tường",
                "TV Bàn", "Tủ đông", "Nền Tủ lạnh", "Tủ lạnh Tường",
                "Nền Máy giặt", "Nền Máy sấy", "Nền Máy rửa chén",
                "KỆ MÁY GIẶT", "Máy giặt Tường", "Máy lạnh Tường",
                "RIG", "Máy nước nóng Tường",
            ],
            "📺 TV Treo Tường (Vị Trí Ưu Tiên)": [
                "TV Treo tường (Vị trí ưu tiên) Vị trí Sony", "Sony (m)",
                "Vị trí Samsung", "Samsung (m)",
                "Vị trí Polytron", "Polytron",
                "Vị trí Sharp", "Sharp",
                "Vị trí Toshiba", "Toshiba",
                "Vị trí TCL", "TCL",
            ],
            "📺 TV Đảo (Island TV)": [
                "TV Đảo (Nguyên tắc) (/Kệ) Samsung",
                "Sharp .1", "Sony", "Polytron .1", "Toshiba .1", "TCL .1",
            ],
            "❌️ Máy Lạnh & Tủ Lạnh Hãng": [
                "Máy lạnh Treo tường (SL) Panasonic",
                "Daikin", "LG", "Samsung",
                "Polytron .2", "Sharp .2", "Midea", "Gree",
                "Aqua", "TCL .2", "Electrolux",
                "Tủ lạnh Treo tường (/Kệ) Midea",
                "TCL .3", "Aqua .1", "Polytron .3", "Sharp .3", "Toshiba .2",
            ],
            "🦺 WM Đảo & SDA": [
                "MÁY GIẶT ĐẢO Midea",
                "TCL .4", "Aqua .2", "Polytron .4", "Sharp .4", "Toshiba .3",
                "TV đầu tiên của dòng", "Kệ SDA đầu tiên",
                "SDA MIYAKO (TƯỜNG)", "MIYAKO (ENDCAP)",
                "PHILIPS (TƯỜNG)", "PHILIPS (ENDCAP)",
                "ELECTROLUX (TƯỜNG)", "ELECTROLUX (ENDCAP)",
                "MIDEA (TƯỜNG)", "MIDEA (ENDCAP)", "Kệ", "Tường.3",
            ],
            "🖼️ Poster & Mặt Tiền & Diện Tích": [
                "TỔNG POSTER TƯỜNG SỬ DỤNG", "CÒN LẠI",
                "POSTER TƯỜNG Thuê theo Thương hiệu",
                "Samsung.1", "Aqua.1", "Polytron.1", "LG.1", "Elux", "Sharp.1",
                "Logo Erablue", "Logo Erafone",
                "Mặt tiền Chính (C)", "Khác (R)", "Khác (L)",
                "Diện tích (m2) Kho Điện máy",
                "WC + Phòng Nhân viên", "Kho + Server", "Bãi đậu xe",
                "Showroom", "Tổng diện tích", "Đất trống",
            ],
        }
    else:
        PARTITIONS = {
            "🌐 All Columns": None,
            "📊 Total Capacity": [
                "Còn lại Bàn", "Tường",
                "Tài nguyên Layout Bàn", "Tường.1",
                "Tài nguyên Thực tế Bàn", "Tường.2",
            ],
            "📱 ICT – Table & Wall (All brands)": [
                "Samsung Bàn Demo", "Tủ Tường Thương hiệu",
                "Apple 1.2m Bàn Demo", "Tủ Tường Thương hiệu.1",
                "OPPO Bàn Demo", "Tủ Tường Thương hiệu.2",
                "Xiaomi Bàn Demo", "Tủ Tường Thương hiệu.3",
                "Vivo Bàn Demo", "Tủ Tường Thương hiệu.4",
                "Realme Bàn Demo", "Tủ Tường Thương hiệu.5",
                "Đa thương hiệu (Huawei, Realme, Infinix) Demo ĐA THƯƠNG HIỆU ( Infinix )",
                "GHI CHÚ",
            ],
            "⚡ Erablue Electronics Resources": [
                "Tài nguyên cho Erablue Electronics TV Treo tường",
                "TV Bàn", "Tủ đông", "Nền Tủ lạnh", "Tủ lạnh Tường",
                "Nền Máy giặt", "Nền Máy sấy", "Nền Máy rửa chén",
                "KỆ MÁY GIẶT", "Máy giặt Tường", "Máy lạnh Tường",
                "RIG", "Máy nước nóng Tường",
            ],
            "📺 Brand TV – Wall (Priority Loc)": [
                "TV Treo tường (Vị trí ưu tiên) Vị trí Sony", "Sony (m)",
                "Vị trí Samsung", "Samsung (m)",
                "Vị trí Polytron", "Polytron",
                "Vị trí Sharp", "Sharp",
                "Vị trí Toshiba", "Toshiba",
                "Vị trí TCL", "TCL",
            ],
            "📺 Island TV": [
                "TV Đảo (Nguyên tắc) (/Kệ) Samsung",
                "Sharp .1", "Sony", "Polytron .1", "Toshiba .1", "TCL .1",
            ],
            "❌️ AC & Fridge by Brand": [
                "Máy lạnh Treo tường (SL) Panasonic",
                "Daikin", "LG", "Samsung",
                "Polytron .2", "Sharp .2", "Midea", "Gree",
                "Aqua", "TCL .2", "Electrolux",
                "Tủ lạnh Treo tường (/Kệ) Midea",
                "TCL .3", "Aqua .1", "Polytron .3", "Sharp .3", "Toshiba .2",
            ],
            "🦺 WM Island & SDA": [
                "MÁY GIẶT ĐẢO Midea",
                "TCL .4", "Aqua .2", "Polytron .4", "Sharp .4", "Toshiba .3",
                "TV đầu tiên của dòng", "Kệ SDA đầu tiên",
                "SDA MIYAKO (TƯỜNG)", "MIYAKO (ENDCAP)",
                "PHILIPS (TƯỜNG)", "PHILIPS (ENDCAP)",
                "ELECTROLUX (TƯỜNG)", "ELECTROLUX (ENDCAP)",
                "MIDEA (TƯỜNG)", "MIDEA (ENDCAP)", "Kệ", "Tường.3",
            ],
            "🖼️ Poster & Facade & Area": [
                "TỔNG POSTER TƯỜNG SỬ DỤNG", "CÒN LẠI",
                "POSTER TƯỜNG Thuê theo Thương hiệu",
                "Samsung.1", "Aqua.1", "Polytron.1", "LG.1", "Elux", "Sharp.1",
                "Logo Erablue", "Logo Erafone",
                "Mặt tiền Chính (C)", "Khác (R)", "Khác (L)",
                "Diện tích (m2) Kho Điện máy",
                "WC + Phòng Nhân viên", "Kho + Server", "Bãi đậu xe",
                "Showroom", "Tổng diện tích", "Đất trống",
            ],
        }

    if sheet_choice == "Erablue Existing":
        part_key = st.selectbox(
            T[lang]["partition_label"],
            list(PARTITIONS.keys()),
            help=T[lang]["partition_help"],
        )
        part_cols = PARTITIONS[part_key]
    else:
        part_key = ""
        part_cols = None

    # Build display dataframe
    CORE_ID_COLS = ["ID Cửa hàng", "Tên Cửa hàng", "Tỉnh/Thành phố (Rút gọn)", "Kích thước Cửa hàng", "Khu vực"]
    if part_cols is not None:
        id_present = [c for c in CORE_ID_COLS if c in filtered.columns]
        extra = [c for c in part_cols if c in filtered.columns and c not in id_present]
        display_df = filtered[id_present + extra]
    else:
        display_df = filtered

    total_rows = len(display_df)

    # Display row count and Export button
    dcol_left, dcol_right = st.columns([3, 1])
    with dcol_left:
        st.markdown(
            f'<div style="padding-top:8px;font-size:13px;color:#64748b;">'
            f'{T[lang]["show_count"].format(count=total_rows, cols=len(display_df.columns))}</div>',
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
    if not display_df.empty:
        html_out = render_sticky_table(display_df, max_height=820)
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

    tab1, tab2 = st.tabs([T[lang]["tab_rek"], T[lang]["tab_fix"]])

    with tab1:
        with st.spinner(T[lang]["loading_sheet"].format(sheet="Reklame store")):
            try:
                df_rek = load_sheet("Reklame store")
            except Exception as e:
                st.error(str(e))
                df_rek = pd.DataFrame()
        if not df_rek.empty:
            scol1, scol2 = st.columns([3, 1])
            with scol1:
                search_rek = st.text_input(T[lang]["search_shop"], key="rek_search")
            if search_rek:
                mask = pd.Series([False] * len(df_rek), index=df_rek.index)
                for col in df_rek.columns[:3]:
                    mask |= df_rek[col].astype(str).str.lower().str.contains(search_rek.lower(), na=False)
                df_rek = df_rek[mask]
            with scol2:
                st.markdown('<div style="height:28px;"></div>', unsafe_allow_html=True)
                xlsx_rek = convert_df_to_excel(df_rek)
                st.download_button(
                    label=T[lang]["export_btn"],
                    data=xlsx_rek,
                    file_name=f"Reklame_store_export_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="rek_export_btn",
                    use_container_width=True
                )
            html_rek = render_sticky_table(df_rek, max_height=700)
            components.html(html_rek, height=750, scrolling=False)
        else:
            st.warning(T[lang]["no_rek"])

    with tab2:
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
                mask = pd.Series([False] * len(df_fix), index=df_fix.index)
                for col in df_fix.columns[:3]:
                    mask |= df_fix[col].astype(str).str.lower().str.contains(search_fix.lower(), na=False)
                df_fix = df_fix[mask]
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

    for col, (label, query) in zip(qcols, presets):
        with col:
            if st.button(label, use_container_width=True):
                st.session_state["ai_input"] = query

    for col, (label, query) in zip(qcols2, presets2):
        with col:
            if st.button(label, use_container_width=True):
                st.session_state["ai_input"] = query

    st.markdown("---")

    # Free-form query
    query_input = st.text_area(
        T[lang]["ai_free_query"],
        value=st.session_state.get("ai_input", ""),
        height=90,
        placeholder=T[lang]["ai_placeholder"],
        key="ai_text_input",
    )

    if st.button(T[lang]["ai_analyze_btn"], type="primary", use_container_width=False):
        if query_input.strip():
            with st.spinner("AI is analyzing live data..."):
                result = analyze(query_input, df_main)
                st.session_state["ai_result"] = result
                st.session_state["ai_query_used"] = query_input
        else:
            st.warning("Vui lòng nhập câu hỏi.")

    # Show result
    if "ai_result" in st.session_state:
        st.markdown("---")
        st.markdown(
            f'<div style="background:#eff6ff;border-left:4px solid #2563eb;'
            f'padding:10px 16px;border-radius:6px;margin-bottom:12px;font-size:13px;">'
            f'{T[lang]["ai_query_label"].format(query=st.session_state.get("ai_query_used",""))}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(st.session_state["ai_result"])

        if st.button(T[lang]["ai_clear_btn"]):
            del st.session_state["ai_result"]
            st.session_state.pop("ai_query_used", None)
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
