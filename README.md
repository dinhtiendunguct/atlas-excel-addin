# Atlas — Excel 365 Windows

Bản mã nguồn đầu tiên cho add-in offline. Chưa đóng gói hoặc chạy thử trong Excel: môi trường tạo mã không khởi động được Excel COM (80070520).

## Tạo và cài add-in

1. Trong Excel: File → Options → Trust Center → Trust Center Settings → Macro Settings → bật tạm **Trust access to the VBA project object model** để script nhập mã nguồn. Không cần chọn bật toàn bộ macro.
2. Mở PowerShell tại thư mục này, chạy `powershell -NoProfile -ExecutionPolicy Bypass -File .\Build-Atlas.ps1`.
3. Khôi phục thiết lập Trust access sau khi tạo xong.
4. Excel → File → Options → Add-ins → Manage Excel Add-ins → Go → Browse → chọn Atlas.xlam. Cho phép macro cho add-in theo chính sách của máy.
5. Mở workbook trống, tab Atlas → Tạo tiến độ. Sửa bảng và bấm Tính lịch và Gantt sau mỗi lần thay đổi.

## Quy ước bản đầu

- ID duy nhất; thời lượng là số nguyên ngày làm việc, 0 là milestone.
- Lịch cố định thứ Hai–thứ Bảy, nghỉ Chủ nhật. Chưa có lịch nghỉ lễ tùy chỉnh.
- Quan hệ nhập `A:FS:0;B:SS:2`, hỗ trợ FS/SS/FF/SF và lag nguyên âm/dương.
- % hoàn thành nhập số 0–100. Thanh xanh lá thể hiện phần hoàn thành, xanh dương là kế hoạch, đỏ là công việc có dự trữ bằng 0.
- Gantt theo ngày, tối đa 731 ngày; chưa có chế độ tuần/tháng hoặc đường nối quan hệ.
- Tính lịch dùng mốc biên ngày làm việc; kết thúc hiển thị ngày làm việc cuối cùng. Milestone hiển thị ở biên bắt đầu tương ứng.
- Lưu baseline chép ngày bắt đầu/kết thúc vào K:L và từ chối ghi đè. Chưa vẽ thanh baseline.
- Chưa có phân cấp WBS, ngày thực tế, remaining duration, rescheduling theo data date, lịch nguồn lực hoặc tự cập nhật khi sửa ô.
- In ngang vừa một trang theo chiều rộng. Cần kiểm tra khổ giấy và độ đọc trước khi xuất báo cáo dài.

## Kiểm tra nghiệm thu trong Excel

Với dữ liệu mẫu: A bắt đầu offset 0, B và C offset 3, D offset 9; dự án 19 ngày làm việc. A/B/D có dự trữ 0, C có dự trữ 2.

Kiểm tra ID trùng, ID thiếu, quan hệ tự phụ thuộc và vòng A→B→A phải báo lỗi trước khi ghi kết quả. Kiểm tra FS/SS/FF/SF có lag, ngày Chủ nhật, milestone và baseline trước khi áp dụng dữ liệu dự án thật.
