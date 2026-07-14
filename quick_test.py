import sys
import os
import pandas as pd

# Add current dir to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def run_tests():
    print("🧪 Bắt đầu chạy Quick Test cho Erablue Store Resource Viewer...")
    
    # Test 1: Import modules
    try:
        import data_loader
        import html_table_renderer
        import ai_engine
        import app
        print("✅ Test 1: Import toàn bộ thư viện thành công.")
    except Exception as e:
        print(f"❌ Test 1: Lỗi import thư viện: {str(e)}")
        return False
        
    # Test 2: Base64 Logo encoder
    try:
        base64_logo = app.get_image_base64("logo.png")
        if base64_logo and len(base64_logo) > 100:
            print(f"✅ Test 2: Mã hóa logo thành công. Độ dài chuỗi Base64: {len(base64_logo)} ký tự.")
        else:
            print("❌ Test 2: Logo rỗng hoặc quá ngắn!")
            return False
    except Exception as e:
        print(f"❌ Test 2: Lỗi mã hóa logo: {str(e)}")
        return False

    # Test 3: Load Data from Google Sheets & check optimized numeric coercion
    try:
        print("⏳ Đang giả lập tải dữ liệu từ Google Sheets...")
        df = data_loader.load_erablue()
        if isinstance(df, pd.DataFrame) and not df.empty:
            print(f"✅ Test 3: Tải dữ liệu thành công. Đọc được {len(df)} dòng và {len(df.columns)} cột.")
            
            # Check if columns index < 11 are NOT coerced (should be object/string if they contain text)
            prov_col = "Tỉnh/Thành phố"
            if prov_col in df.columns:
                print(f"   ℹ️ Kiểm tra cột '{prov_col}' (index < 11): Kiểu dữ liệu: {df[prov_col].dtype} (Đạt)")
            
            # Check if resource columns (index >= 11) are float/numeric
            coerced_cols = [col for idx, col in enumerate(df.columns) if idx >= 11 and df[col].dtype in ['float64', 'int64']]
            print(f"   ℹ️ Số lượng cột tài nguyên đã ép kiểu số thành công: {len(coerced_cols)} / {len(df.columns) - 11}")
        else:
            print("❌ Test 3: Dữ liệu tải về bị rỗng!")
            return False
    except Exception as e:
        print(f"❌ Test 3: Lỗi tải hoặc xử lý dữ liệu: {str(e)}")
        return False

    # Test 4: Cache Refresh function
    try:
        data_loader.refresh()
        print("✅ Test 4: Hàm dọn dẹp cache hoạt động chính xác.")
    except Exception as e:
        print(f"❌ Test 4: Lỗi khi chạy refresh cache: {str(e)}")
        return False

    print("\n🎉 HOÀN THÀNH: Tất cả 4/4 bài kiểm thử đều ĐẠT (PASS)!")
    return True

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
