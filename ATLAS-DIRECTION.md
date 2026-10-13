# Atlas của Đinh Tiến Dũng

Căn cứ: nội dung đọc trực tiếp từ chat **Project Atlas Accelerated**, ID `6a92834d-ac70-83ec-a042-06dc76919cc6`, ngày 13/10/2026. Link người dùng cung cấp: https://chatgpt.com/share/6acdddae-d100-83ec-92e5-924880b27c56 (trình đọc web không tải được; đã đọc chat trong ứng dụng).

Atlas hướng tới công cụ điều hành công trường: tiến độ → điều kiện thi công → năng suất và tổ đội → phương án thu hồi tiến độ → dòng tiền. Ảnh giao diện chỉ dùng tham khảo trình bày.

## Đã triển khai trong 0.4.0

- Lookahead 21 ngày lịch tính từ ngày báo cáo TIEN_DO!B5, bao gồm việc đang làm và việc chưa hoàn thành tồn quá hạn, loại HM và việc đã đạt 100%.
- Liên kết duy nhất bằng Activity ID, không tự ghép theo tên, bộ môn hoặc zone.
- Sáu điều kiện: bản vẽ, vật tư, mặt bằng, QA/QC, HSE, nhân lực. PASS cần người xác nhận, ngày không sau ngày báo cáo, tham chiếu bằng chứng.
- Ràng buộc cần ID, Activity ID, chủ trì, hạn xử lý; CLOSED cần người xác nhận, ngày đóng hợp lệ, bằng chứng. Chưa đủ vẫn HOLD.
- ID mồ côi và dữ liệu trùng chặn READY; không suy diễn thiếu thông tin thành sẵn sàng.
- READY biểu thị điều kiện đã được người dùng xác nhận trong sổ, chưa phải phê duyệt huy động hoặc chi tiền. Phần mềm không xác thực nội dung tài liệu chứng cứ.
- Bổ sung sheet cho workbook 0.3 mà không sửa/xóa dữ liệu tiến độ, baseline, lịch sử kỳ cũ.

## Các bước tiếp theo, chưa triển khai

- Khối lượng còn lại, nhật ký sản lượng / giờ công, năng suất đo được.
- Dự báo nhu cầu nhân lực 21 ngày, năng lực nghề dùng chung, giới hạn mặt bằng zone.
- So sánh phương án tăng tốc, tính lại tác động CPM lên milestone, không cộng ngày rút ngắn cục bộ thành ngày thu hồi dự án.
- Liên kết chi phí được phát hành / có điều kiện / nghĩa vụ đã cam kết; thiếu dữ liệu Finance giữ trống, không mặc định 0.
- Quy trình phê duyệt, kiểm tra nguồn bằng chứng và nhật ký thay đổi xác nhận.

Không nhập số liệu mô phỏng từ chat H2 thành số liệu hiện trường thật. Các trạng thái toolkit/test trong chat là thông tin lịch sử, chưa được kiểm tra lại trên Google Sheets trong lần phát triển này.
