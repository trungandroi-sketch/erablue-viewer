"""
html_table_renderer.py – Renders a DataFrame as a premium HTML table with:
  • Frozen "ID Cửa hàng" + "Tên Cửa hàng" columns (sticky left, always visible)
  • Two-row grouped & color-coded headers (Group label + column name)
  • Summary row at the bottom (SUM + store-count per column)

Column names match the actual Vietnamese Google Sheet headers.
"""
import html as _html
import pandas as pd

# ─── Helper: strip leading/trailing whitespace and newlines from col names ────
def _clean(col: str) -> str:
    return " ".join(col.split())


# ─── Column group definitions (Vietnamese column names from Google Sheets) ────
# Columns are stripped of whitespace before matching.
COLUMN_GROUPS = [
    {
        "label": "Thông Tin Cửa Hàng", "emoji": "📋",
        "color": "#1e3a5f", "text": "#ffffff",
        "cols": [
            "ID Cửa hàng", "Tên Cửa hàng", "Địa chỉ", "Quận/Huyện",
            "Tỉnh/Thành phố", "Tỉnh/Thành phố (Rút gọn)", "Khu vực",
            "AM", "RSM / Khu vực", "Ghi chú & Ngoại lệ",
            "Trưng bày đặc biệt", "Kích thước Cửa hàng", "Ngày Setup", "ƯỚC TÍNH GO",
        ],
    },
    {
        "label": "Công Suất Tổng", "emoji": "📊",
        "color": "#0f4c81", "text": "#ffffff",
        "cols": [
            "Còn lại Bàn", "Tường", "Tài nguyên Layout Bàn", "Tường.1",
            "Tài nguyên Thực tế Bàn", "Tường.2",
        ],
    },
    {
        "label": "Samsung", "emoji": "📱",
        "color": "#1428a0", "text": "#ffffff",
        "cols": ["Samsung Bàn Demo", "Tủ Tường Thương hiệu"],
    },
    {
        "label": "Apple", "emoji": "🍎",
        "color": "#3a3a3c", "text": "#ffffff",
        "cols": ["Apple 1.2m Bàn Demo", "Tủ Tường Thương hiệu.1"],
    },
    {
        "label": "OPPO", "emoji": "📱",
        "color": "#1a7c40", "text": "#ffffff",
        "cols": ["OPPO Bàn Demo", "Tủ Tường Thương hiệu.2"],
    },
    {
        "label": "Xiaomi", "emoji": "📱",
        "color": "#e05c00", "text": "#ffffff",
        "cols": ["Xiaomi Bàn Demo", "Tủ Tường Thương hiệu.3"],
    },
    {
        "label": "Vivo", "emoji": "📱",
        "color": "#3251ff", "text": "#ffffff",
        "cols": ["Vivo Bàn Demo", "Tủ Tường Thương hiệu.4"],
    },
    {
        "label": "Realme", "emoji": "📱",
        "color": "#c49000", "text": "#ffffff",
        "cols": ["Realme Bàn Demo", "Tủ Tường Thương hiệu.5"],
    },
    {
        "label": "Đa Thương Hiệu", "emoji": "📱",
        "color": "#52606d", "text": "#ffffff",
        "cols": [
            "Đa thương hiệu (Huawei, Realme, Infinix) Demo ĐA THƯƠNG HIỆU ( Infinix )",
            "GHI CHÚ",
        ],
    },
    {
        "label": "Phụ Kiện & Laptop", "emoji": "💻",
        "color": "#374151", "text": "#ffffff",
        "cols": [
            "Tủ Điện thoại Điện thoại", "Máy tính bảng",
            "LAPTOP PRINCIPLE HP", "INTEL", "ACER", "ASUS", "LENOVO",
            "Laptop Bàn Laptop", "Tủ Laptop",
            "Phụ kiện Tường", "Fixture (Sàn)", "CCTV + Loa",
        ],
    },
    {
        "label": "Tài Nguyên Erablue", "emoji": "⚡",
        "color": "#064e3b", "text": "#ffffff",
        "cols": [
            "Tài nguyên cho Erablue Electronics TV Treo tường",
            "TV Bàn", "Tủ đông", "Nền Tủ lạnh", "Tủ lạnh Tường",
            "Nền Máy giặt", "Nền Máy sấy", "Nền Máy rửa chén",
            "KỆ MÁY GIẶT", "Máy giặt Tường", "Máy lạnh Tường",
            "RIG", "Máy nước nóng Tường",
        ],
    },
    {
        "label": "TV Treo Tường (Vị Trí)", "emoji": "📺",
        "color": "#6d28d9", "text": "#ffffff",
        "cols": [
            "TV Treo tường (Vị trí ưu tiên) Vị trí Sony", "Sony (m)",
            "Vị trí Samsung", "Samsung (m)",
            "Vị trí Polytron", "Polytron",
            "Vị trí Sharp", "Sharp",
            "Vị trí Toshiba", "Toshiba",
            "Vị trí TCL", "TCL",
        ],
    },
    {
        "label": "TV Đảo (Island)", "emoji": "📺",
        "color": "#4c1d95", "text": "#ffffff",
        "cols": [
            "TV đảo principle (Nguyên tắc) (/Kệ) Samsung",
            "Sharp .1", "Sony", "Polytron .1", "Toshiba .1", "TCL .1",
        ],
    },
    {
        "label": "Máy Lạnh Treo Tường", "emoji": "❄️",
        "color": "#0369a1", "text": "#ffffff",
        "cols": [
            "Máy lạnh principle (SL) Panasonic",
            "Daikin", "LG", "Samsung",
            "Polytron .2", "Sharp .2", "Midea", "Gree",
            "Aqua", "TCL .2", "Electrolux",
        ],
    },
    {
        "label": "Tủ Lạnh Treo Tường", "emoji": "🧊",
        "color": "#0284c7", "text": "#ffffff",
        "cols": [
            "Tủ lạnh principle (/Kệ) Midea",
            "TCL .3", "Aqua .1", "Polytron .3", "Sharp .3", "Toshiba .2",
        ],
    },
    {
        "label": "Máy Giặt Đảo (Island WM)", "emoji": "🧺",
        "color": "#0f766e", "text": "#ffffff",
        "cols": [
            "MÁY GIAT PRINCIPLE Midea",
            "TCL .4", "Aqua .2", "Polytron .4", "Sharp .4", "Toshiba .3",
        ],
    },
    {
        "label": "SDA", "emoji": "🔌",
        "color": "#6b21a8", "text": "#ffffff",
        "cols": [
            "TV đầu tiên của dòng", "Kệ SDA đầu tiên",
            "SDA MIYAKO (TƯỜNG)", "MIYAKO (ENDCAP)",
            "PHILIPS (TƯỜNG)", "PHILIPS (ENDCAP)",
            "ELECTROLUX (TƯỜNG)", "ELECTROLUX (ENDCAP)",
            "MIDEA (TƯỜNG)", "MIDEA (ENDCAP)", "Kệ", "Tường.3",
        ],
    },
    {
        "label": "Poster Tường", "emoji": "🖼️",
        "color": "#92400e", "text": "#ffffff",
        "cols": [
            "TỔNG POSTER TƯỜNG SỬ DỤNG", "CÒN LẠI",
            "POSTER TƯỜNG Thuê theo Thương hiệu",
            "Samsung.1", "Aqua.1", "Polytron.1", "LG.1", "Elux", "Sharp.1",
            "Logo Erablue", "Logo Erafone",
            "1 đổi 1", "Giao hàng và lắp đặt miễn phí",
            "Tổng đài", "Website", "INTEL.1",
        ],
    },
    {
        "label": "Bàn & Mặt Tiền & Diện Tích", "emoji": "📐",
        "color": "#78350f", "text": "#ffffff",
        "cols": [
            "Bàn BÀN TƯ VẤN",
            "Mặt tiền Chính (C)", "Khác (R)", "Khác (L)",
            "Diện tích (m2) Kho Điện máy",
            "WC + Phòng Nhân viên", "Kho + Server", "Bãi đậu xe",
            "Showroom", "Tổng diện tích", "Đất trống",
        ],
    },
]

FROZEN_COLS = ["ID Cửa hàng", "Tên Cửa hàng", "Tỉnh/Thành phố (Rút gọn)"]
FROZEN_WIDTHS = {"ID Cửa hàng": 108, "Tên Cửa hàng": 240, "Tỉnh/Thành phố (Rút gọn)": 140}

_FALLBACK_GROUP = {"label": "Khác", "emoji": "📄", "color": "#374151", "text": "#ffffff"}

GROUP_TRANSLATIONS = {
    "Samsung": "Samsung",
    "Apple": "Apple",
    "OPPO": "OPPO",
    "Xiaomi": "Xiaomi",
    "Vivo": "Vivo",
    "Realme": "Realme",
    "Đa Thương Hiệu": "Multi-brand",
    "Phụ Kiện & Laptop": "Accessories & Laptop",
    "Tài Nguyên Erablue": "Erablue Resources",
    "TV Treo Tường (Vị Trí)": "TV Wall (Positions)",
    "TV Đảo (Island)": "TV Island",
    "Máy Lạnh Treo Tường": "Wall AC",
    "Tủ Lạnh Treo Tường": "Wall Fridge",
    "Máy Giặt Đảo (Island WM)": "Island Washing Machine",
    "SDA": "SDA",
    "Poster Tường": "Poster Wall",
    "Bàn & Mặt Tiền & Diện Tích": "Table & Facade & Area",
}

COLUMN_TRANSLATIONS = {
    "ID Cửa hàng": "Store ID",
    "ID Store": "Store ID",
    "Store ID": "Store ID",
    "Tên Cửa hàng": "Store Name",
    "Shop Name": "Store Name",
    "Store Name": "Store Name",
    "Tỉnh/Thành phố (Rút gọn)": "Province (Short)",
    "Province (Short)": "Province (Short)",
    "Kích thước Cửa hàng": "Store Size",
    "Size": "Size",
    "Khu vực": "Region",
    "Region": "Region",
    "Quận/Huyện": "District",
    "Địa chỉ": "Address",
    "AM": "AM",
    "RSM / Khu vực": "RSM / Region",
    "ƯỚC TÍNH GO": "GO Estimate",
    "Ngày Setup": "Setup Date",
    "Bàn BÀN TƯ VẤN": "Consulting Table",
    "Mặt tiền Chính (C)": "Main Facade (C)",
    "Khác (R)": "Other (R)",
    "Khác (L)": "Other (L)",
    "Diện tích (m2) Kho Điện máy": "Warehouse Area (m2)",
    "WC + Phòng Nhân viên": "WC & Staff Room",
    "Kho + Server": "Warehouse & Server",
    "Bãi đậu xe": "Parking Area",
    "Showroom": "Showroom Area",
    "Tổng diện tích": "Total Area",
    "Đất trống": "Empty Land",
    "Vách": "Wall Brand",
    "GHI CHÚ": "Notes",
    "Kệ": "SDA Shelves",
    "Tường.3": "SDA Wall",
}


# ─── Helpers ──────────────────────────────────────────────────────────────────
def _normalize(col: str) -> str:
    """Normalize column name: collapse whitespace and newlines."""
    return " ".join(col.split())


def _col_to_group_map(all_cols: list[str]) -> dict[str, dict]:
    """Build a mapping from actual (normalized) column names → group definition."""
    # Normalize defined group cols
    grp_lookup: dict[str, dict] = {}
    for g in COLUMN_GROUPS:
        for gc in g["cols"]:
            grp_lookup[_normalize(gc)] = g

    result = {}
    for c in all_cols:
        nc = _normalize(c)
        result[c] = grp_lookup.get(nc)
    return result


def _short(col: str) -> str:
    """Return a short display name by removing common group prefix patterns."""
    nc = _normalize(col)
    if "Tủ Tường Thương hiệu" in nc:
        return "Vách"
    if "Đa thương hiệu" in nc:
        return "Đa thương hiệu"
    STRIP = [
        "Samsung ", "Apple 1.2m ", "OPPO ", "Xiaomi ", "Vivo ", "Realme ",
        "Tài nguyên cho Erablue Electronics ", "Tài nguyên Layout ",
        "Tài nguyên Thực tế ",
        "TV Treo tường (Vị trí ưu tiên) ",
        "TV Đảo (Nguyên tắc) (/Kệ) ",
        "Máy lạnh Treo tường (SL) ",
        "Tủ lạnh Treo tường (/Kệ) ",
        "Tủ lạnh principle (/Kệ) ",
        "MÁY GIẶT ĐẢO ",
        "MÁY GIAT PRINCIPLE ",
        "SDA ", "POSTER TƯỜNG ",
        "Tủ Tường Thương hiệu",
        "LAPTOP PRINCIPLE ",
        "Diện tích (m2) ",
        "Bàn ",
    ]
    for prefix in STRIP:
        if nc.startswith(prefix):
            remainder = nc[len(prefix):]
            return remainder if remainder else nc
    return nc


def _fmt(val) -> tuple[str, str]:
    """Return (display_string, css_class). cls: 'p'=positive, 'z'=zero, ''=text."""
    try:
        if pd.isna(val):
            return "", "z"
        f = float(val)
        if f == 0:
            return "", "z"
        if f > 0:
            s = str(int(f)) if f == int(f) else f"{f:.1f}"
            return s, "p"
        return str(f), ""
    except (TypeError, ValueError):
        sv = "" if pd.isna(val) else str(val).strip()
        return sv, ""


# ─── Main renderer ────────────────────────────────────────────────────────────
def render_sticky_table(df: pd.DataFrame, max_height: int = 820, lang: str = "vi") -> str:
    """
    Returns a self-contained HTML page string.
    Render with: st.components.v1.html(html, height=max_height+40, scrolling=False)
    """
    df = df.copy()
    # Drop internal id column
    for drop_c in ["id", "Unnamed: 139", "Unnamed: 140", ". "]:
        if drop_c in df.columns:
            df = df.drop(columns=[drop_c])

    col_grp = _col_to_group_map(list(df.columns))

    # Re-order: frozen first
    frozen_present = [c for c in FROZEN_COLS if c in df.columns]
    non_frozen = [c for c in df.columns if c not in frozen_present]
    all_cols = frozen_present + non_frozen
    df = df[[c for c in all_cols if c in df.columns]]
    all_cols = list(df.columns)

    # Compute sticky left offsets
    offsets: dict[str, int] = {}
    x = 0
    for c in all_cols:
        if c in frozen_present:
            offsets[c] = x
            x += FROZEN_WIDTHS.get(c, 150)

    # Build group spans
    spans = []
    i = 0
    while i < len(all_cols):
        c = all_cols[i]
        if c in frozen_present:
            spans.append({"frozen": True, "col": c})
            i += 1
        else:
            g = col_grp.get(c)
            key = g["label"] if g else f"__{c}__"
            grp_cols = [c]
            j = i + 1
            while j < len(all_cols) and all_cols[j] not in frozen_present:
                ng = col_grp.get(all_cols[j])
                nk = ng["label"] if ng else f"__{all_cols[j]}__"
                if nk == key:
                    grp_cols.append(all_cols[j])
                    j += 1
                else:
                    break
            spans.append({"frozen": False, "group": g, "cols": grp_cols})
            i = j

    # ── Header row 1 (group labels) ───────────────────────────────────────
    h1 = ""
    for s in spans:
        if s["frozen"]:
            c = s["col"]
            lft = offsets[c]
            w = FROZEN_WIDTHS.get(c, 150)
            label = _short(c)
            if lang == "en" and label in COLUMN_TRANSLATIONS:
                label = COLUMN_TRANSLATIONS[label]
            h1 += (
                f'<th rowspan="2" style="position:sticky;left:{lft}px;top:0;z-index:25;'
                f'min-width:{w}px;max-width:{w}px;background:#0f2744;color:#93c5fd;'
                f'font-weight:700;padding:8px 10px;white-space:nowrap;text-align:left;'
                f'border-right:2px solid rgba(255,255,255,.2);font-size:11px;">{_html.escape(label)}</th>'
            )
        else:
            g = s.get("group") or _FALLBACK_GROUP
            n = len(s["cols"])
            grp_label = g["label"]
            if lang == "en" and grp_label in GROUP_TRANSLATIONS:
                grp_label = GROUP_TRANSLATIONS[grp_label]
            lbl = _html.escape(f"{g['emoji']} {grp_label}")
            h1 += (
                f'<th colspan="{n}" style="position:sticky;top:0;z-index:12;'
                f'background:{g["color"]};color:{g["text"]};font-size:10px;font-weight:700;'
                f'letter-spacing:.5px;text-transform:uppercase;padding:7px 12px;'
                f'text-align:center;white-space:nowrap;'
                f'border-left:1px solid rgba(255,255,255,.15);">{lbl}</th>'
            )

    # ── Header row 2 (individual column names) ────────────────────────────
    h2 = ""
    for s in spans:
        if s["frozen"]:
            continue
        g = s.get("group") or _FALLBACK_GROUP
        for c in s["cols"]:
            sn = _short(c)
            if lang == "en" and sn in COLUMN_TRANSLATIONS:
                sn = COLUMN_TRANSLATIONS[sn]
            h2 += (
                f'<th style="position:sticky;top:33px;z-index:11;'
                f'background:{g["color"]};filter:brightness(.78);color:{g["text"]};'
                f'font-size:10px;font-weight:600;padding:5px 8px;text-align:center;'
                f'white-space:nowrap;min-width:82px;'
                f'border-left:1px solid rgba(255,255,255,.1);">{_html.escape(sn)}</th>'
            )

    # ── Body rows ─────────────────────────────────────────────────────────
    body = ""
    for ri, (_, row) in enumerate(df.iterrows()):
        bg = "#ffffff" if ri % 2 == 0 else "#f8fafd"
        body += (
            f'<tr style="background:{bg};" '
            f'onmouseover="this.style.background=\'#dbeafe\'" '
            f'onmouseout="this.style.background=\'{bg}\'">'
        )
        for c in all_cols:
            val = row[c]
            disp, cls = _fmt(val)
            if c in frozen_present:
                lft = offsets[c]
                w = FROZEN_WIDTHS.get(c, 150)
                body += (
                    f'<td style="position:sticky;left:{lft}px;z-index:3;'
                    f'min-width:{w}px;max-width:{w+20}px;background:#f0f5ff;'
                    f'border-right:2px solid #bfdbfe;color:#1e3a5f;font-weight:600;'
                    f'padding:4px 10px;white-space:nowrap;overflow:hidden;'
                    f'text-overflow:ellipsis;text-align:left;">{_html.escape(disp)}</td>'
                )
            else:
                color_map = {
                    "p": "color:#15803d;font-weight:600;",
                    "z": "color:#d1d5db;",
                    "": "",
                }
                cs = color_map.get(cls, "")
                body += (
                    f'<td style="padding:4px 8px;text-align:center;white-space:nowrap;'
                    f'border-left:1px solid #f1f5f9;font-size:12px;{cs}">'
                    f'{_html.escape(disp)}</td>'
                )
        body += "</tr>"

    # ── Summary row ───────────────────────────────────────────────────────
    sum_cells = ""
    for c in all_cols:
        if c == FROZEN_COLS[0] if FROZEN_COLS else None:
            lft = offsets[c]
            w = FROZEN_WIDTHS[c]
            total_lbl = "∑ TOTAL" if lang == "en" else "∑ TỔNG"
            sum_cells += (
                f'<td style="position:sticky;left:{lft}px;z-index:3;min-width:{w}px;'
                f'background:#0a1f44;color:#60a5fa;font-weight:700;padding:6px 10px;'
                f'text-align:left;border-top:2px solid #3b82f6;">{total_lbl}</td>'
            )
        elif len(frozen_present) > 1 and c == FROZEN_COLS[1]:
            lft = offsets[c]
            w = FROZEN_WIDTHS[c]
            store_count_lbl = f"{len(df)} stores" if lang == "en" else f"{len(df)} cửa hàng"
            sum_cells += (
                f'<td style="position:sticky;left:{lft}px;z-index:3;min-width:{w}px;'
                f'background:#0a1f44;color:#60a5fa;font-weight:700;padding:6px 10px;'
                f'text-align:left;border-top:2px solid #3b82f6;">{store_count_lbl}</td>'
            )
        elif c in frozen_present:
            lft = offsets[c]
            w = FROZEN_WIDTHS.get(c, 120)
            sum_cells += (
                f'<td style="position:sticky;left:{lft}px;z-index:3;min-width:{w}px;'
                f'background:#0a1f44;border-top:2px solid #3b82f6;"></td>'
            )
        else:
            try:
                num = pd.to_numeric(df[c], errors="coerce")
                tot = num.sum()
                cnt = int((num > 0).sum())
                if cnt > 0:
                    ts = str(int(tot)) if tot == int(tot) else f"{tot:.1f}"
                    sum_cells += (
                        f'<td style="background:#1e3a5f;color:#93c5fd;font-weight:700;'
                        f'font-size:10px;text-align:center;padding:4px 6px;'
                        f'border-top:2px solid #3b82f6;white-space:nowrap;">'
                        f'∑{ts}<br><small style="opacity:.7">({cnt}✓)</small></td>'
                    )
                else:
                    sum_cells += '<td style="background:#1e3a5f;border-top:2px solid #3b82f6;"></td>'
            except Exception:
                sum_cells += '<td style="background:#1e3a5f;border-top:2px solid #3b82f6;"></td>'

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  *{{box-sizing:border-box;margin:0;padding:0;}}
  html,body{{font-family:'Inter',sans-serif;font-size:12px;background:transparent;overflow:hidden;}}
  .wrap{{overflow:auto;max-height:{max_height}px;border:1px solid #e2e8f0;
    border-radius:10px;box-shadow:0 4px 20px rgba(30,58,95,.10);background:#fff;}}
  table{{border-collapse:separate;border-spacing:0;min-width:max-content;}}
  tbody tr td{{border-bottom:1px solid #f1f5f9;}}
</style>
</head>
<body>
<div class="wrap">
<table>
<thead>
<tr>{h1}</tr>
<tr>{h2}</tr>
</thead>
<tbody>
{body}
<tr>{sum_cells}</tr>
</tbody>
</table>
</div>
<script>
(function(){{
  var r1=document.querySelector('thead tr:first-child');
  if(r1){{
    var h=r1.getBoundingClientRect().height;
    document.querySelectorAll('thead tr:nth-child(2) th').forEach(function(t){{t.style.top=h+'px';}});
  }}
}})();
</script>
</body>
</html>"""
