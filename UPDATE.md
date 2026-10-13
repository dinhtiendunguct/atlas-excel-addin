# Cập nhật Atlas

## Trạng thái

Đã thêm bộ cập nhật vào mã nguồn 0.2.0. File Atlas.xlam hiện có là bản cũ: cần build lại một lần để có nút mới. Chưa có repository/URL phát hành nên feedUrl đang trống; chưa thể tải phiên bản online thực tế.

## Cài một lần

Lưu và đóng Excel. Từ thư mục project chạy:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\0. Chat GPT\Projects\add in tiến độ\Build-Atlas.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\0. Chat GPT\Projects\add in tiến độ\Install-Atlas.ps1"
```

Trong danh sách Excel Add-ins, bỏ chọn bản Atlas tại thư mục cũ. Browse đến `%LOCALAPPDATA%\AtlasExcel\Atlas.xlam` và chọn bản đó. Chỉ nạp một bản Atlas.

Build vẫn cần quyền Trust access to VBA project object model như trước; cập nhật các phiên bản sau không cần quyền này.

## Phát hành

Chủ dự án cấu hình `feedUrl` trong update-config.json bằng HTTPS URL cố định của latest.json. Đối với GitHub Releases có thể dùng `https://github.com/OWNER/REPO/releases/latest/download/latest.json`; public repository không cần token. Không hỗ trợ đăng nhập vào private repository trong bản đầu.

Build add-in rồi tạo manifest:

```powershell
.\New-AtlasRelease.ps1 -Version 0.2.1 -DownloadUrl 'https://github.com/OWNER/REPO/releases/download/v0.2.1/Atlas.xlam' -Notes 'Mo ta thay doi'
```

Upload Atlas.xlam và latest.json vào cùng release. Version phải tăng. SHA256 trong manifest phải khớp đúng file đã upload. Nguồn phát hành phải do bạn quản lý: SHA256 phát hiện file lỗi/thay đổi so với manifest, không thay thế chữ ký số của nhà phát hành.

## Sử dụng

Atlas → Kiểm tra cập nhật → đồng ý tải → lưu và đóng tất cả Excel khi được nhắc. Công cụ chờ tối đa 10 phút, không ép tắt Excel. Sau thành công mở lại Excel bằng tay.

Bản cũ được giữ tại Atlas.previous-*.xlam. Nếu cần khôi phục: đóng Excel, giữ bản hiện tại ở tên khác rồi sao chép bản previous mong muốn thành Atlas.xlam; sửa version trong update-config.json theo bản khôi phục.

Không thay workbook dự án, không đổi cài đặt bảo mật Excel. Không tự chuyển đổi cấu trúc workbook. Bộ cập nhật PowerShell là bootstrap cố định; cơ chế này hiện cập nhật Atlas.xlam, chưa tự cập nhật chính bootstrap.

## Kiểm tra

Đã kiểm tra cú pháp PowerShell và gắn callback Ribbon ở mã nguồn. Chưa xác nhận chạy trong Excel và chưa thử tải end-to-end từ máy chủ phát hành. Cần nghiệm thu: không có bản mới, từ chối tải, sai hash, mất mạng, Excel chưa đóng, cập nhật thành công và khôi phục.
