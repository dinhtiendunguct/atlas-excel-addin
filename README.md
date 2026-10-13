# Atlas Excel Add-in 0.3.0

Add-in lập tiến độ thi công chạy offline trên Excel Microsoft 365 Windows, với giao diện tiếng Việt tham khảo bố cục Ribbon / bảng WBS / Gantt và báo cáo nguồn lực.

## Cài đặt

Tải `Atlas-Setup-0.3.0.zip` tại [GitHub Releases](https://github.com/dinhtiendunguct/atlas-excel-addin/releases/latest). Giải nén, lưu và đóng toàn bộ Excel, rồi nhấp đúp `Install-Atlas.cmd`.

Bộ cài chép file vào `%LOCALAPPDATA%\AtlasExcel`, bỏ chặn các file của Atlas và đăng ký add-in với Excel. Nó không thay đổi thiết lập macro hoặc Protected View của Excel. Nếu chính sách máy chặn đăng ký, bộ cài hiển thị đường dẫn để Browse thủ công.

Mở Excel → Atlas → **Mẫu demo** để xem dự án 5 tầng + 1 hầm, hoặc **Tạo dự án** để tạo workbook dữ liệu mới. Lưu workbook dạng `.xlsx` riêng với chương trình.

## Bản 0.3 có gì

- Ribbon: Dự án, Công việc, Tính toán, Gantt, Nguồn lực & báo cáo, Hỗ trợ.
- 7 sheet: TIEN_DO, NGUON_LUC, CAU_HINH, NHAT_KY, HUONG_DAN, BIEU_DO_NL, BAO_CAO.
- Hạng mục HM, công việc CV, mốc MOC; phân cấp / đánh số WBS và thu gọn hạng mục.
- Lập lịch CPM với FS/SS/FF/SF và lag nguyên âm/dương; phát hiện ID trùng, ID thiếu, tự phụ thuộc và quan hệ vòng.
- Lịch tuần tùy chỉnh và ngày nghỉ bổ sung.
- Gantt ngày / tuần / tháng, đường găng, % hoàn thành, milestone, ngày báo cáo, baseline và bật/tắt đường nối.
- Lưu baseline một lần; hiển thị sai lệch kết thúc bằng ngày làm việc.
- Gán nhân lực, thiết bị và đơn giá; biểu đồ bình quân tháng và chi phí tháng / lũy kế.
- Ghi nhận % thực tế theo kỳ, không ghi đè lịch sử cùng ngày; đường cong S kế hoạch và các điểm thực tế đã ghi nhận.

## Quy ước dữ liệu

`A01:FS:0;A02:SS:2` trong cột Công việc trước; ID giữ ổn định khi đánh lại WBS. Hạng mục HM tổng hợp công việc con theo tiền tố WBS và không tham gia mạng CPM.

`NC:12;KS:1` trong cột Nguồn lực. Số lượng cố định mỗi ngày làm việc của công việc. Đơn giá / ngày do người dùng nhập tại NGUON_LUC. Các đơn giá mặc định và dữ liệu demo chỉ để minh họa.

Lịch tuần gồm 7 ký tự từ thứ Hai đến Chủ nhật: `0000001` nghỉ Chủ nhật; `0000011` nghỉ thứ Bảy và Chủ nhật. `0` là ngày làm việc, `1` là ngày nghỉ. Ngày nghỉ bổ sung nhập tại CAU_HINH cột A từ dòng 12.

## Giới hạn cần biết

Bản này tính lịch từ ngày khởi công; chưa có cập nhật mạng theo actual dates / remaining duration, cân bằng nguồn lực, nhiều calendar riêng cho activity hoặc phê duyệt OBS. Chỉnh dữ liệu rồi bấm Cập nhật tiến độ / Báo cáo; chưa tự tính sau mỗi lần sửa ô.

% tổng thể có trọng số thời lượng CV, không tính trùng HM và không tính mốc. Đường cong S kế hoạch phân bổ tuyến tính theo ngày làm việc, dùng baseline khi toàn bộ CV có baseline; nếu chưa đầy đủ thì dùng lịch hiện tại. Chỉ số TT/KH theo thời lượng không phải SPI của EVM. Thực tế chỉ có các kỳ được ghi nhận, không tự tạo lịch sử quá khứ.

Tối đa 2000 dòng, 3650 ngày làm việc; Gantt ngày tối đa 731 ngày. Gantt in mặc định A3 ngang; dự án dài nên chọn thang tháng hoặc chỉnh vùng in để bảo đảm đọc được.

Workbook từ bản 0.2 được giữ nguyên, chưa chuyển đổi tự động. Tạo dự án 0.3 trong workbook mới trước khi nhập dữ liệu.

## Kiểm tra

Đã chạy 30 ca nghiệm thu trực tiếp trong Excel 16.0: FS/SS/FF/SF, merge network, float, ngày nghỉ lễ / tuần 5 ngày, milestone, baseline, ID và quan hệ lỗi, Gantt ngày/tháng, đường nối, snapshot và đối chiếu chi phí nguồn lực. Đã kiểm tra bản in Gantt và các báo cáo. Chưa nghiệm thu bộ cài tự đăng ký trên máy khác hoặc luồng tải cập nhật ở mọi cấu hình bảo mật.

## Phát triển

`Build-Atlas.ps1` đóng gói Atlas.Core.bas, Atlas.UI.bas, Atlas.Reports.bas và Atlas.Tests.bas. Unicode literal được chuyển sang mã hex để giữ dấu tiếng Việt khi VBA editor dùng ANSI. Build cần Excel và quyền Trust access to the VBA project object model; cài và cập nhật không cần quyền đó.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Build-Atlas.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\Test-Atlas.ps1 -Visible
```

Chi tiết phát hành tại [UPDATE.md](UPDATE.md).
