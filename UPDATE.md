# Cập nhật Atlas

Nguồn phát hành công khai: https://github.com/dinhtiendunguct/atlas-excel-addin

## Người dùng

Nếu đã có bộ cập nhật, chọn **Atlas → Kiểm tra cập nhật**. Đồng ý tải, lưu workbook, đóng tất cả Excel theo thông báo. Công cụ chờ tối đa 10 phút, không ép tắt Excel; mở lại Excel sau khi cập nhật thành công.

Nếu chưa cài bootstrap hoặc cần cài lại: tải ZIP mới nhất tại Releases, giải nén, đóng Excel và nhấp đúp `Install-Atlas.cmd`. Bộ cài sao lưu Atlas đã có, chép bootstrap và cấu hình mới, tự đăng ký add-in. Nó không đổi chính sách bảo mật Excel. Một số máy do tổ chức quản lý có thể yêu cầu Browse thủ công hoặc chặn VBA theo chính sách riêng.

Chỉ nạp một bản Atlas. Add-in được cài cố định tại `%LOCALAPPDATA%\AtlasExcel\Atlas.xlam`. Bộ cài bỏ chọn bản Atlas.xlam ở đường dẫn khác khi đăng ký bản mới.

## Phát hành phiên bản

Build và chạy Test-Atlas.ps1 trước. Tăng version trong update-config.json. Tạo manifest:

```powershell
.\New-AtlasRelease.ps1 -Version 0.3.0 -DownloadUrl 'https://github.com/dinhtiendunguct/atlas-excel-addin/releases/download/v0.3.0/Atlas.xlam' -Notes 'Mô tả thay đổi'
```

Upload đúng Atlas.xlam và latest.json vào release. Feed cố định:

`https://github.com/dinhtiendunguct/atlas-excel-addin/releases/latest/download/latest.json`

Tải qua HTTPS, so sánh version, kiểm tra SHA256 và cấu trúc VBA trong ZIP trước khi thay thế. SHA256 đối chiếu với manifest không thay thế chữ ký số của nhà phát hành.

Online updater hiện chỉ thay Atlas.xlam; bootstrap PowerShell giữ nguyên. Thay bootstrap cần bộ cài ZIP mới. Workbook dự án không bị sửa khi cập nhật chương trình; bản 0.2 chưa tự chuyển cấu trúc sang 0.3.

## Khôi phục

Atlas.previous-*.xlam được lưu cạnh bản đang dùng. Đóng Excel, giữ bản hiện tại bằng tên khác, sao chép bản previous mong muốn thành Atlas.xlam và sửa version trong update-config.json theo phiên bản khôi phục.
