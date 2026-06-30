"""
ai_engine.py – Natural language query engine for Erablue Viewer.
Works directly with pandas DataFrames using actual Vietnamese Google Sheet column names.
"""
from __future__ import annotations
import re
import pandas as pd

# ─── Normalize column name (collapse whitespace/newlines) ────────────────────
def _n(col: str) -> str:
    return " ".join(col.split())


# ─── Brand definitions (Vietnamese column names from Google Sheets) ───────────
BRANDS = [
    {
        "kw": ["samsung"],
        "name": "Samsung",
        "color": "#1428a0",
        "table": "Samsung Bàn Demo",
        "wall": "Tủ Tường Thương hiệu",        # index 0
    },
    {
        "kw": ["apple", "iphone", "ipad"],
        "name": "Apple",
        "color": "#3a3a3c",
        "table": "Apple 1.2m Bàn Demo",
        "wall": "Tủ Tường Thương hiệu.1",
    },
    {
        "kw": ["oppo"],
        "name": "OPPO",
        "color": "#1a7c40",
        "table": "OPPO Bàn Demo",
        "wall": "Tủ Tường Thương hiệu.2",
    },
    {
        "kw": ["xiaomi", "redmi"],
        "name": "Xiaomi",
        "color": "#e05c00",
        "table": "Xiaomi Bàn Demo",
        "wall": "Tủ Tường Thương hiệu.3",
    },
    {
        "kw": ["vivo"],
        "name": "Vivo",
        "color": "#3251ff",
        "table": "Vivo Bàn Demo",
        "wall": "Tủ Tường Thương hiệu.4",
    },
    {
        "kw": ["realme"],
        "name": "Realme",
        "color": "#c49000",
        "table": "Realme Bàn Demo",
        "wall": "Tủ Tường Thương hiệu.5",
    },
    {
        "kw": ["infinix", "huawei", "multibrand", "đa thương hiệu"],
        "name": "Multibrand (Infinix/Huawei)",
        "color": "#52606d",
        "table": "Đa thương hiệu (Huawei, Realme, Infinix) Demo ĐA THƯƠNG HIỆU ( Infinix )",
        "wall": None,
    },
]

# Location columns (normalized)
LOCATION_COLS_NORM = [
    "Tỉnh/Thành phố (Rút gọn)",
    "Khu vực",
    "Quận/Huyện",
    "Địa chỉ",
    "Tỉnh/Thành phố",
]

# Area column
AREA_COL = "Khu vực"
ID_COL = "ID Cửa hàng"
NAME_COL = "Tên Cửa hàng"
PROV_COL = "Tỉnh/Thành phố (Rút gọn)"
SIZE_COL = "Kích thước Cửa hàng"


# ─── Helpers ──────────────────────────────────────────────────────────────────
def _normalize_df_cols(df: pd.DataFrame) -> dict[str, str]:
    """Return mapping normalized_name → actual_column_name."""
    return {_n(c): c for c in df.columns}


def _find_col(df: pd.DataFrame, normalized_name: str) -> str | None:
    """Find actual column in df by normalized name."""
    for c in df.columns:
        if _n(c) == _n(normalized_name):
            return c
    return None


def _count_positive(df: pd.DataFrame, col_normalized: str) -> tuple[int, pd.DataFrame]:
    """Count rows where column > 0."""
    col = _find_col(df, col_normalized)
    if col is None:
        return 0, pd.DataFrame()
    num = pd.to_numeric(df[col], errors="coerce").fillna(0)
    mask = num > 0
    return int(mask.sum()), df[mask].copy()


def _tbl(rows: pd.DataFrame, n: int = 25) -> str:
    id_c = _find_col(rows, ID_COL)
    nm_c = _find_col(rows, NAME_COL)
    ar_c = _find_col(rows, AREA_COL)
    pv_c = _find_col(rows, PROV_COL)
    sz_c = _find_col(rows, SIZE_COL)

    lines = ["| ID | Tên Cửa Hàng | Khu Vực | Tỉnh/TP | Size |",
             "|---|---|---|---|---|"]
    for _, r in rows.head(n).iterrows():
        lines.append(
            f"| {r.get(id_c,'') if id_c else ''} "
            f"| {r.get(nm_c,'') if nm_c else ''} "
            f"| {r.get(ar_c,'') if ar_c else ''} "
            f"| {r.get(pv_c,'') if pv_c else ''} "
            f"| {r.get(sz_c,'') if sz_c else ''} |"
        )
    rest = len(rows) - n
    if rest > 0:
        lines.append(f"| ... | *+{rest} cửa hàng khác* | | | |")
    return "\n".join(lines)


# ─── Filters Extractor ────────────────────────────────────────────────────────
def _apply_query_filters(q: str, df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    df_filtered = df.copy()
    filters_applied = []
    
    # 0. Store ID/Name filter
    id_col = _find_col(df, ID_COL)
    if id_col:
        # Find all numbers in the query
        nums = re.findall(r'\b\d+\b', q)
        for num in nums:
            match_rows = df[df[id_col].astype(str).str.strip() == num]
            if not match_rows.empty:
                df_filtered = match_rows
                name_col = _find_col(df, NAME_COL)
                name_str = f" ({match_rows.iloc[0][name_col]})" if name_col else ""
                filters_applied.append(f"Cửa hàng ID: **{num}**{name_str}")
                return df_filtered, filters_applied

    # 1. Size filter
    size_val = None
    # Matches 'size s', 'cỡ s', 'loại s', 'quy mô s', etc.
    m = re.search(r'\b(?:size|cỡ|loại|quy mô)\s+(s|m|l|xl|xxl)\b', q)
    if m:
        size_val = m.group(1).upper()
    else:
        # Match standalone 's', 'm', 'l', 'xl', 'xxl' preceded by 'shop' or 'cửa hàng'
        m2 = re.search(r'\b(?:cửa hàng|shop)\s+(s|m|l|xl|xxl)\b', q)
        if m2:
            size_val = m2.group(1).upper()
            
    if size_val:
        size_col = None
        for c in ["Kích thước Cửa hàng", "Size", "Kích thước"]:
            size_col = _find_col(df, c)
            if size_col:
                break
        if size_col:
            df_filtered = df_filtered[df_filtered[size_col].astype(str).str.strip().str.upper() == size_val]
            filters_applied.append(f"Size: **{size_val}**")
            
    # 2. Location filter
    loc_kw = [
        "banten", "tangerang", "jakarta", "depok", "bekasi",
        "bogor", "karawang", "bandung", "garut", "subang",
        "ciledug", "jawa", "jkr", "jbr", "cikarang",
        "jawa barat", "jawa tengah", "jawa timur", "yogyakarta",
        "dki jakarta", "tây java", "trung java", "đông java"
    ]
    
    found_locs = []
    for loc in loc_kw:
        if re.search(rf'\b{loc}\b', q):
            found_locs.append(loc)
            
    if found_locs:
        best_loc = max(found_locs, key=len)
        masks = []
        for col_norm in LOCATION_COLS_NORM:
            actual = _find_col(df, col_norm)
            if actual:
                masks.append(df_filtered[actual].astype(str).str.lower().str.contains(best_loc, na=False))
        if masks:
            combined = masks[0]
            for m in masks[1:]:
                combined |= m
            df_filtered = df_filtered[combined]
            filters_applied.append(f"Khu vực/Tỉnh thành: **{best_loc.title()}**")
            
    return df_filtered, filters_applied


# ─── Brand-Category Resolver ──────────────────────────────────────────────────
def _resolve_specific_column(q: str, df: pd.DataFrame) -> str | None:
    brand_map = {
        "sony": "Sony",
        "samsung": "Samsung",
        "polytron": "Polytron",
        "sharp": "Sharp",
        "toshiba": "Toshiba",
        "tcl": "TCL",
        "panasonic": "Panasonic",
        "daikin": "Daikin",
        "lg": "LG",
        "midea": "Midea",
        "gree": "Gree",
        "aqua": "Aqua",
        "electrolux": "Electrolux",
        "elux": "Electrolux",
    }
    
    found_brand = None
    for kw, brand in brand_map.items():
        if re.search(rf'\b{kw}\b', q):
            found_brand = brand
            break
            
    if not found_brand:
        return None
        
    is_ac = any(w in q for w in ["máy lạnh", "ac", "điều hòa", "air conditioner"])
    is_fridge = any(w in q for w in ["tủ lạnh", "fridge", "refrigerator"])
    is_wm = any(w in q for w in ["máy giặt", "washing machine", "wm"])
    is_island_tv = any(w in q for w in ["tv đảo", "tivi đảo", "island tv", "kệ tv"])
    is_wall_tv = any(w in q for w in ["tv treo", "tivi treo", "treo tường", "wall tv"]) or (any(w in q for w in ["tivi", "tv"]) and not is_island_tv and not is_ac and not is_fridge and not is_wm)

    if is_wall_tv:
        candidates = []
        for col in df.columns:
            col_l = col.lower()
            if found_brand.lower() in col_l:
                if not re.search(r'\.\d+', col):
                    candidates.append(col)
        if candidates:
            candidates.sort(key=lambda c: ("vị trí" not in c.lower(), len(c)))
            return candidates[0]

    elif is_island_tv:
        for col in df.columns:
            col_l = col.lower()
            if found_brand.lower() in col_l:
                if ".1" in col or "đảo" in col_l:
                    return col
        for col in df.columns:
            if f"{found_brand} .1" in col or f"{found_brand}.1" in col:
                return col

    elif is_ac:
        for col in df.columns:
            col_l = col.lower()
            if found_brand.lower() in col_l:
                if ".2" in col or "máy lạnh" in col_l:
                    return col
        for col in df.columns:
            if f"{found_brand} .2" in col or f"{found_brand}.2" in col:
                return col

    elif is_fridge:
        for col in df.columns:
            col_l = col.lower()
            if "tủ lạnh" in col_l and found_brand.lower() in col_l:
                return col
        suffixes = {"Midea": "Midea", "TCL": ".3", "Aqua": ".1", "Polytron": ".3", "Sharp": ".3", "Toshiba": ".2"}
        target_sfx = suffixes.get(found_brand, "")
        for col in df.columns:
            col_l = col.lower()
            if found_brand.lower() in col_l and target_sfx.lower() in col_l:
                return col

    elif is_wm:
        for col in df.columns:
            col_l = col.lower()
            if "máy giặt" in col_l and found_brand.lower() in col_l:
                return col
        suffixes = {"Midea": "Midea", "TCL": ".4", "Aqua": ".2", "Polytron": ".4", "Sharp": ".4", "Toshiba": ".3"}
        target_sfx = suffixes.get(found_brand, "")
        for col in df.columns:
            col_l = col.lower()
            if found_brand.lower() in col_l and target_sfx.lower() in col_l:
                return col

    return None


# ─── Single Store Report Renderer ─────────────────────────────────────────────
def _render_single_store_view(df_filtered: pd.DataFrame, col_name: str, kind_desc: str, filter_header: str, df_original: pd.DataFrame) -> str:
    store_row = df_filtered.iloc[0]
    id_col = _find_col(df_original, ID_COL)
    name_col = _find_col(df_original, NAME_COL)
    area_col = _find_col(df_original, AREA_COL)
    prov_col = _find_col(df_original, PROV_COL)
    size_col = _find_col(df_original, SIZE_COL)
    
    store_id = store_row.get(id_col, "") if id_col else ""
    store_name = store_row.get(name_col, "") if name_col else ""
    val = store_row.get(col_name)
    
    # Check if positive
    is_positive = False
    try:
        is_positive = float(val) > 0
    except Exception:
        is_positive = pd.notna(val) and str(val).strip() not in ["", "-", "0", "0.0"]
        
    status_emoji = "✅ CÓ" if is_positive else "❌ KHÔNG"
    status_text = "có triển khai / sở hữu" if is_positive else "không triển khai / không sở hữu"
    
    val_str = f" (Giá trị ghi nhận: `{val}`)" if pd.notna(val) else ""
    
    return f"""### 🏪 Kết quả tra cứu cửa hàng **{store_name}** (ID: `{store_id}`)
{filter_header}
> **{status_emoji}** – Cửa hàng này **{status_text}** tài nguyên **`{kind_desc}`**{val_str}.

---

#### 📋 Thông tin chi tiết của cửa hàng:
- **Khu vực**: {store_row.get(area_col, '') if area_col else ''}
- **Tỉnh/Thành phố (Rút gọn)**: {store_row.get(prov_col, '') if prov_col else ''}
- **Kích thước (Size)**: {store_row.get(size_col, '') if size_col else ''}
- **Giá trị tài nguyên `{col_name}`**: `{val}`

_💡 Bạn có thể tra cứu các tài nguyên khác của cửa hàng này bằng cách nhập tên tài nguyên (vd: "cửa hàng {store_id} có máy lạnh Daikin không")._
"""


# ─── Main query function ──────────────────────────────────────────────────────
def analyze(query: str, df: pd.DataFrame) -> str:
    """Parse a Vietnamese NL query and return a markdown report."""
    q = query.lower().strip()
    total_original = len(df)
    if total_original == 0:
        return "⚠️ Không có dữ liệu để phân tích."

    # 1. Apply high-precision filters first (Store ID, Size, Location)
    df_filtered, filters_applied = _apply_query_filters(q, df)
    total = len(df_filtered)
    
    filter_header = ""
    if filters_applied:
        filter_header = f"> 📌 **Bộ lọc tự động áp dụng**: {', '.join(filters_applied)}\n>\n"

    # 2. Try Google Gemini API with the FILTERED data!
    import os
    import streamlit as st
    
    gemini_key = None
    try:
        if "GEMINI_API_KEY" in st.secrets:
            gemini_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
        
    if not gemini_key:
        gemini_key = os.environ.get("GEMINI_API_KEY")
        
    if gemini_key:
        try:
            import google.generativeai as genai
            genai.configure(api_key=gemini_key)
            
            # Self-healing model selection based on API key permissions
            available_models = []
            try:
                for m in genai.list_models():
                    if "generateContent" in m.supported_generation_methods:
                        available_models.append(m.name)
            except Exception:
                pass
                
            model_name = "gemini-2.5-flash"
            if available_models:
                preferences = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro", "gemini-pro"]
                matched = None
                for pref in preferences:
                    for am in available_models:
                        if pref in am:
                            matched = am
                            break
                    if matched:
                        break
                if matched:
                    model_name = matched
                else:
                    model_name = available_models[0]
            
            model_clean = model_name.replace("models/", "")
            model = genai.GenerativeModel(model_clean)
            
            # Dynamically build model display name
            model_parts = model_clean.split("-")
            if len(model_parts) >= 3:
                model_display_name = f"Gemini {model_parts[1]} {model_parts[2].title()}"
            elif len(model_parts) == 2:
                model_display_name = f"Gemini {model_parts[1].title()}"
            else:
                model_display_name = f"Gemini {model_clean.title()}"
            
            # Optimize DataFrame size: keep only active columns of df_filtered!
            cols_to_keep = []
            core_cols = [ID_COL, NAME_COL, SIZE_COL, AREA_COL, PROV_COL, "Ngày Setup", "ƯỚC TÍNH GO"]
            for col in df_filtered.columns:
                if col in core_cols:
                    cols_to_keep.append(col)
                    continue
                series_str = df_filtered[col].astype(str).str.strip().str.lower()
                is_active = series_str.apply(lambda x: x not in ["", "-", "0", "0.0", "nan", "none"]).any()
                if is_active:
                    cols_to_keep.append(col)
            
            df_compact = df_filtered[cols_to_keep].fillna("")
            csv_data = df_compact.to_csv(index=False)
            
            prompt = (
                "Bạn là trợ lý AI thông minh phân tích dữ liệu cửa hàng Erablue Electronics.\n"
                "Dưới đây là bảng dữ liệu thực tế từ Google Sheet chứa thông tin các cửa hàng, tài nguyên trưng bày (bàn demo, vách thương hiệu các hãng Samsung, Apple, OPPO, Xiaomi, Vivo, Realme, Đa thương hiệu), thiết bị điện máy (tivi treo, tivi đảo, máy lạnh, tủ lạnh, máy giặt), và các cột thông tin phụ trợ (Ngày Setup, ƯỚC TÍNH GO, Kích thước Cửa hàng, Địa chỉ, Khu vực, Tỉnh/Thành phố (Rút gọn)).\n"
                "Nhiệm vụ của bạn là phân tích dữ liệu này và trả lời câu hỏi của người dùng một cách chính xác.\n\n"
                "QUY TẮC TRẢ LỜI:\n"
                "1. Trả lời bằng ngôn ngữ tương thích với câu hỏi của người dùng (Ưu tiên Tiếng Việt hoặc Tiếng Anh theo ngôn ngữ của câu hỏi), định dạng Markdown đẹp mắt.\n"
                "2. Nếu câu hỏi yêu cầu thống kê (ví dụ: đếm số shop, tỷ lệ %), hãy tính toán chính xác dựa trên dữ liệu.\n"
                "3. Nếu câu hỏi liên quan đến một cửa hàng cụ thể (bằng ID số hoặc Tên shop), hãy hiển thị thông tin dạng bảng hoặc thẻ thông tin chi tiết các thuộc tính liên quan đến câu hỏi.\n"
                "4. Hãy viết câu trả lời trực tiếp, không lặp lại câu hỏi của người dùng.\n"
                "5. Dữ liệu Google Sheets chứa thông tin thực tế, không tự tiện bịa đặt các dữ liệu nằm ngoài bảng CSV dưới đây.\n"
                "6. Lưu ý: Cột 'Bãi đậu xe' biểu thị diện tích sân xe/bãi đỗ xe của cửa hàng. Hãy đọc kỹ giá trị tương ứng ở cột này.\n\n"
                f"BẢNG DỮ LIỆU CSV ({len(df_filtered)} CỬA HÀNG):\n"
                "```csv\n"
                f"{csv_data}\n"
                "```\n\n"
                f"CÂU HỎI CỦA NGƯỜI DÙNG: {query}"
            )
            
            response = model.generate_content(
                prompt,
                generation_config={"temperature": 0.2}
            )
            
            if response.text:
                return f"""### 🤖 Phân Tích Bởi Siêu Trí Tuệ Nhân Tạo ({model_display_name})
{filter_header}
{response.text}

---
_💡 Phản hồi này được sinh ra bởi mô hình **Google {model_display_name}** dựa trên dữ liệu thực tế từ Google Sheets._
"""
        except Exception as e:
            st.warning(f"⚠️ Trợ lý AI (Gemini API) gặp lỗi hoặc chưa cấu hình key: {e}. Hệ thống tự động chuyển sang công cụ phân tích cục bộ.")
            pass

    # ── 0. High-precision brand-category resolver ───────────────────────────
    resolved_col = _resolve_specific_column(q, df)
    if resolved_col:
        count, sub = _count_positive(df_filtered, resolved_col)
        pct = count / total * 100 if total else 0
        
        if total == 1:
            return _render_single_store_view(df_filtered, resolved_col, resolved_col, filter_header, df)
            
        return f"""### 📊 Kết quả phân tích cho câu hỏi của bạn
{filter_header}
> Tìm thấy cột phù hợp nhất trong bảng dữ liệu: **`{resolved_col}`**
>
> Ghi nhận **{count}** / **{total}** cửa hàng có giá trị hoạt động (>0) ở cột này.
> → Tỷ lệ: **{pct:.1f}%**  {'🟢' if pct >= 50 else '🟡' if pct >= 25 else '🔴'}

---

#### 🏪 Danh sách cửa hàng:

{_tbl(sub, 25)}
"""

    # ── 1. Brand ICT queries ────────────────────────────────────────────────
    for brand in BRANDS:
        if any(kw in q for kw in brand["kw"]):
            want_wall = any(w in q for w in ["vách", "vach", "wall", "tường", "tủ tường", "cabinet"])
            col_name = brand["wall"] if (want_wall and brand["wall"]) else brand["table"]
            kind = "Vách / Tủ Tường" if (want_wall and brand["wall"]) else "Bàn Demo (Table)"

            count, sub = _count_positive(df_filtered, col_name)
            pct = count / total * 100 if total else 0

            if total == 1:
                return _render_single_store_view(df_filtered, col_name, f"{brand['name']} {kind}", filter_header, df)

            return f"""### 📊 Kết quả – **{brand["name"]}** ({kind})
{filter_header}
> Hệ thống ghi nhận **{count}** / **{total}** cửa hàng có {kind} của nhãn **{brand["name"]}**
> → Tỷ lệ phủ: **{pct:.1f}%**  {'🟢' if pct >= 50 else '🟡' if pct >= 25 else '🔴'}

---

#### 🏪 Danh sách {count} cửa hàng có {kind} **{brand["name"]}**:

{_tbl(sub, 25)}

---

#### 💡 Insight
- **{total - count}** cửa hàng chưa có {kind.lower()} {brand["name"]} → tiềm năng mở rộng.
- Độ phủ {pct:.1f}% {'tốt ✅' if pct >= 50 else 'trung bình ⚠️' if pct >= 25 else 'thấp – cần cải thiện 🔴'}.
"""

    # ── 2. Location queries ─────────────────────────────────────────────────
    loc_kw = [
        "banten", "tangerang", "jakarta", "depok", "bekasi",
        "bogor", "karawang", "bandung", "garut", "subang",
        "ciledug", "jawa", "jkr", "jbr", "cikarang",
    ]
    for loc in loc_kw:
        if loc in q and not any(w in q for w in ["tất cả", "hãng", "brand", "tổng", "khu vực", "phân bổ", "độ phủ", "laptop", "tv", "tivi", "ac", "máy lạnh", "tủ lạnh", "máy giặt"]):
            masks = []
            for col_norm in LOCATION_COLS_NORM:
                actual = _find_col(df_filtered, col_norm)
                if actual:
                    masks.append(df_filtered[actual].astype(str).str.lower().str.contains(loc, na=False))
            if masks:
                combined = masks[0]
                for m in masks[1:]:
                    combined |= m
                sub = df_filtered[combined].copy()
                count = len(sub)
                pct = count / total_original * 100
                return f"""### 📍 Cửa hàng tại **{loc.title()}**
{filter_header}
> Tìm thấy **{count}** / **{total_original}** cửa hàng ({pct:.1f}%) tại khu vực **'{loc.title()}'**.

---

{_tbl(sub, 30)}
"""

    # ── 3. Area distribution ────────────────────────────────────────────────
    if any(w in q for w in ["khu vực", "area", "thống kê", "phân bổ", "vùng", "phân phối"]):
        area_col = _find_col(df_filtered, AREA_COL)
        if area_col:
            by_area = df_filtered.groupby(area_col).size().sort_values(ascending=False)
            lines = ["| Khu Vực | Số Cửa Hàng | Tỷ Lệ |", "|---|---|---|"]
            for area, cnt in by_area.items():
                lines.append(f"| {area} | {cnt} | {cnt/total*100:.1f}% |" if total else f"| {area} | {cnt} | 0.0% |")
            return f"""### 🗺️ Phân Bổ Cửa Hàng Theo Khu Vực
{filter_header}
> Tổng **{total}** cửa hàng · **{len(by_area)}** khu vực

{"  " + chr(10).join(lines)}
"""

    # ── 4. All brands summary ───────────────────────────────────────────────
    if any(w in q for w in ["tổng", "tất cả", "hãng", "độ phủ", "coverage", "brand", "tổng hợp"]):
        lines = ["| Thương Hiệu | Bàn Demo | Vách/Tủ Tường | Tỷ Lệ Bàn |", "|---|---|---|---|"]
        for brand in BRANDS[:6]:
            tc, _ = _count_positive(df_filtered, brand["table"])
            wc = 0
            if brand["wall"]:
                wc, _ = _count_positive(df_filtered, brand["wall"])
            pct = tc / total * 100 if total else 0
            bar = "█" * int(pct // 5) + "░" * (20 - int(pct // 5))
            lines.append(f"| **{brand['name']}** | {tc} | {wc} | {pct:.1f}% `{bar}` |")
        return f"""### 📊 Tổng Hợp Độ Phủ Tất Cả Thương Hiệu ICT
{filter_header}
> Hệ thống: **{total}** cửa hàng

{"  " + chr(10).join(lines)}

_💡 Hỏi thêm: "Có mấy shop có bàn OPPO?" / "Shop nào ở Banten?"_
"""

    # ── 5. Dynamic Column Search fallback ─────────────────────────────────────
    # Extract keywords by splitting query and removing common stop words
    words = q.replace("?", "").replace(".", "").replace(",", "").split()
    stop_words = {
        "có", "mấy", "shop", "cửa", "hàng", "nào", "ở", "tổng", "số", "lượng",
        "danh", "sách", "tìm", "cho", "tôi", "biết", "với", "hỏi", "được", "không",
        "các", "những", "cái", "chiếc", "của", "nhãn", "hãng", "hiệu", "thương",
        "size", "cỡ", "loại", "quy", "mô", "s", "m", "l", "xl", "xxl"
    }
    keywords = [w for w in words if w not in stop_words and len(w) > 1]
    
    translations = {
        "tv": ["tv", "television", "tivi"],
        "tivi": ["tv", "television", "tivi"],
        "laptop": ["laptop", "máy tính xách tay"],
        "fridge": ["fridge", "tủ lạnh", "refrigerator"],
        "tủ lạnh": ["fridge", "tủ lạnh", "refrigerator"],
        "mesin cuci": ["mesin cuci", "máy giặt", "wm"],
        "máy giặt": ["mesin cuci", "máy giặt", "wm"],
        "air conditioner": ["ac", "máy lạnh", "máy điều hòa"],
        "máy lạnh": ["ac", "máy lạnh", "máy điều hòa"],
        "ac": ["ac", "máy lạnh", "máy điều hòa"],
        "water heater": ["nước nóng", "water heater"],
        "nước nóng": ["nước nóng", "water heater"],
        "bàn": ["bàn", "table"],
        "vách": ["vách", "wall", "tường", "cabinet"],
        "tường": ["vách", "wall", "tường", "cabinet"],
        "sân xe": ["bãi đậu xe", "sân xe", "bãi xe", "đậu xe", "đỗ xe"],
        "đậu xe": ["bãi đậu xe", "sân xe", "bãi xe", "đậu xe", "đỗ xe"],
        "đỗ xe": ["bãi đậu xe", "sân xe", "bãi xe", "đậu xe", "đỗ xe"],
        "bãi xe": ["bãi đậu xe", "sân xe", "bãi xe", "đậu xe", "đỗ xe"],
        "bãi đậu xe": ["bãi đậu xe", "sân xe", "bãi xe", "đậu xe", "đỗ xe"],
    }
    
    # Expand keywords using translations
    expanded_kws = []
    for kw in keywords:
        expanded_kws.append(kw)
        if kw in translations:
            expanded_kws.extend(translations[kw])
            
    # Try to find matching columns
    matched_cols = []
    for col in df_filtered.columns:
        col_lower = col.lower()
        score = sum(1 for kw in keywords if kw in col_lower)
        for kw in keywords:
            if kw in translations:
                if any(t_kw in col_lower for t_kw in translations[kw]):
                    score += 1
        if score > 0:
            matched_cols.append((col, score))
            
    if matched_cols:
        matched_cols.sort(key=lambda x: -x[1])
        best_score = matched_cols[0][1]
        best_cols = [col for col, score in matched_cols if score == best_score]
        
        # Filter for the first matched column to display stats
        best_col = best_cols[0]
        count, sub = _count_positive(df_filtered, best_col)
        pct = count / total * 100 if total else 0
        
        if total == 1:
            return _render_single_store_view(df_filtered, best_col, best_col, filter_header, df)

        other_matches_str = ""
        if len(best_cols) > 1:
            other_matches_str = "\n*Các cột liên quan khác tìm thấy: " + ", ".join(f"`{c}`" for c in best_cols[1:]) + "*"
            
        return f"""### 📊 Kết quả phân tích cho câu hỏi của bạn
{filter_header}
> Tìm thấy cột phù hợp nhất trong bảng dữ liệu: **`{best_col}`**{other_matches_str}
>
> Ghi nhận **{count}** / **{total}** cửa hàng có giá trị hoạt động (>0) ở cột này.
> → Tỷ lệ: **{pct:.1f}%**  {'🟢' if pct >= 50 else '🟡' if pct >= 25 else '🔴'}

---

#### 🏪 Danh sách cửa hàng:

{_tbl(sub, 25)}
"""

    # ── 6. Default fallback ─────────────────────────────────────────────────
    lines = []
    for brand in BRANDS[:6]:
        cnt, _ = _count_positive(df_filtered, brand["table"])
        lines.append(f"- **{brand['name']}**: {cnt} cửa hàng có bàn demo ({cnt/total*100:.1f}%)" if total else f"- **{brand['name']}**: {cnt} cửa hàng (0.0%)")

    return f"""### 📊 Tổng Quan Hệ Thống Erablue
{filter_header}
Hệ thống đang quản lý **{total}** cửa hàng.
 
#### Độ phủ bàn demo ICT:
{"  " + chr(10).join(lines)}
 
---
#### 💬 Thử hỏi:
- "Có mấy shop có bàn OPPO?"
- "Có mấy shop có vách Samsung?"
- "Shop nào ở Banten?"
- "Có mấy shop có bàn laptop?"
- "Tổng hợp độ phủ tất cả hãng?"
- "Phân bổ cửa hàng theo khu vực?"
"""
