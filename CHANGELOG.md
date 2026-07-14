# Changelog

## [2026-07-14]
### Added
- Thêm file test tự động `quick_test.py` để verify hoạt động tải dữ liệu, mã hóa logo, và dọn cache.
- Tích hợp **Logo Erablue Deployment** chính thức hình tròn vào Sidebar.
- Thêm thuộc tính `will-change` kích hoạt card đồ họa (GPU) để render animation chuyển tab mượt mà.

### Changed
- Nâng cấp toàn diện giao diện sang phong cách **Creative Tim** với thiết kế kính mờ (Glassmorphism), dải màu Tím Hoàng gia và thẻ tiêu đề nền trắng.
- Tăng kích thước font chữ toàn trang lên **`15px`** giúp tăng độ rõ nét và dễ đọc.
- Tối ưu hóa hiệu năng ép kiểu số: Chỉ ép kiểu cho các cột tài nguyên (từ cột 11 trở đi), giúp giảm tải CPU khi load bảng.

### Fixed
- Sửa lỗi nghiêm trọng **Đóng băng bộ nhớ đệm** (`_XLSX_BYTES` cache stale) giúp hệ thống tự động làm mới dữ liệu sau mỗi 5 phút.
- Sửa lỗi **Màu chữ nút bấm bị ẩn** (trùng màu nền trắng) ở Sidebar.
- Sửa lỗi **NameError: name 'box' is not defined** do thiếu thoát ngoặc nhọn `{}` trong CSS của bảng chi tiết.
- Bọc lỗi biên dịch AI Query trong khối `try...except` để đảm bảo ứng dụng không bao giờ bị sập giao diện.
