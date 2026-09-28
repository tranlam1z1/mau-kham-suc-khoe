# Prompt tạo trang web "Phiếu kiểm tra sức khỏe học sinh"

Dán nguyên văn prompt dưới đây vào Claude Code (hoặc công cụ AI lập trình khác), đính kèm file Word mẫu `kham-suc-khoe-hoc-sinh.docx`.

---

Hãy tạo một trang web **một file duy nhất `index.html`** (HTML + CSS + JavaScript thuần, không cần build, mở trực tiếp bằng trình duyệt) để người dùng điền **"Phụ lục III – Nội dung kiểm tra sức khỏe đầu năm học cho học sinh trong các cơ sở giáo dục phổ thông"** theo đúng file Word mẫu đính kèm, và **xuất ra file Word (.docx) giữ đúng bố cục mẫu**.

## 1. Giao diện nhập liệu (tiếng Việt, gọn, dùng tốt trên điện thoại)
Chia thành các thẻ (card) theo đúng thứ tự mẫu:

**THÔNG TIN HÀNH CHÍNH**
1. Họ và tên (tự chuyển CHỮ IN HOA)
2. Giới tính: radio Nam / Nữ
3. Ngày sinh (date) + Tuổi (tự tính từ ngày sinh, cho phép sửa tay)
4. Số CC/CCCD/Hộ chiếu/Mã định danh
5. Cấp ngày (date), Nơi cấp
6. Dân tộc (có gợi ý: Kinh, Tày, Thái, Mường, Khmer, Hoa, Nùng, H'Mông, Dao) · 7. Đối tượng
8. Nguồn chi trả · 9. Nhóm máu (gợi ý A, B, AB, O)
10. Nơi ở hiện tại: Tỉnh/thành, Phường/xã, Số nhà/thôn/xóm; Tên cơ sở giáo dục đang theo học
11. Lý do khám sức khỏe

**TIỀN SỬ BỆNH TẬT**
- 1. Tiền sử gia đình: radio Không/Có; nếu "Có" hiện ô ghi tên bệnh
- 2a. Sản khoa: radio Bình thường / Không bình thường; nếu "Không bình thường" hiện checkbox: Đẻ thiếu tháng, Đẻ thừa cân, Đẻ có can thiệp, Đẻ ngạt, Mẹ bị bệnh trong thời kỳ mang thai + ô ghi tên bệnh
- 2b. Tiêm chủng: bảng 7 dòng (BCG; Bạch hầu, ho gà, uốn ván; Sởi; Bại liệt; Viêm não Nhật Bản B; Viêm gan B; Các loại khác – có ô ghi tên), mỗi dòng radio Có / Không / Không nhớ rõ
- 2c. Tiền sử bệnh/tật (bẩm sinh, mãn tính): Không/Có + ô ghi tên bệnh
- 2d. Đang điều trị bệnh: Không/Có + ô ghi tên bệnh và thuốc đang dùng

**I. KHÁM THỂ LỰC**: Chiều cao (cm), Cân nặng (kg), BMI (tự tính = kg / m², 1 chữ số thập phân), Mạch (lần/phút), Huyết áp tâm thu / tâm trương (mmHg)

**II. KHÁM LÂM SÀNG**
1. Mắt: thị lực không kính (phải/trái), có kính (phải/trái), các bệnh về mắt
2. Tai–Mũi–Họng: tai trái/phải nói thường & nói thầm (m), bất thường TMH
3. Răng–Hàm–Mặt: hàm trên, hàm dưới, bất thường khác
4. Cơ xương, thần kinh, tâm thần: Vận động không đối xứng (Không/Có); Lưng, cột sống (Bình thường/Không bình thường); Dáng đi (Bình thường/Không bình thường); Buồn bã, lo âu, mất ngủ kéo dài (Có/Không); Hành vi bất thường hoặc nguy cơ tự gây hại (Có/Không); ô "Nếu có đề nghị ghi rõ"

**III. KẾT LUẬN CHUNG**: Tình trạng sức khỏe; Cần tư vấn/chuyển khám chuyên khoa; Địa danh, Ngày ký, Họ tên người kết luận

## 2. Xuất file Word
- Dùng thư viện `docx` qua CDN: `https://cdn.jsdelivr.net/npm/docx@9.6.1/dist/index.iife.js`, tạo file bằng `docx.Packer.toBlob`.
- Khổ giấy Letter, lề 1 inch, font Arial 10pt như mẫu.
- Tiêu đề căn giữa, in đậm: "Phụ lục III", "NỘI DUNG KIỂM TRA SỨC KHỎE ĐẦU NĂM HỌC CHO HỌC SINH", "TRONG CÁC CƠ SỞ GIÁO DỤC PHỔ THÔNG", dòng nghiêng "(Ban hành kèm theo Thông tư số 33/2026/TT-BYT ngày 29 tháng 8 năm 2026 của Bộ trưởng Bộ Y tế)".
- Giữ nguyên câu chữ, thứ tự, đánh số của mẫu. Ô được chọn in `☒`, ô không chọn in `☐`.
- Ô trống thì in dấu chấm "……" như mẫu (dùng tab có dấu chấm dẫn tới lề phải cho các dòng dài); ô đã điền thì in giá trị.
- Bảng tiêm chủng 5 cột (STT | Loại vắc xin | Có | Không | Không nhớ rõ), hàng tiêu đề gộp ô "Tình trạng tiêm/uống vắc xin", lặp tiêu đề khi sang trang.
- Mục II đặt trong bảng 1 cột có viền, dòng tiêu đề từng mục (1. Mắt, 2. Tai - Mũi - Họng, 3. Răng - Hàm - Mặt) tách riêng như mẫu; không cắt ngang ô khi sang trang.
- Khối ký bên phải, không viền: "…, ngày … tháng … năm …", "NGƯỜI KẾT LUẬN", "(Ký, ghi rõ họ tên và đóng dấu)", chừa chỗ ký, rồi họ tên.
- Tên file: `KSK_<HO_TEN_KHONG_DAU>.docx`.

## 3. Tiện ích
- Thanh nút cố định dưới cùng: **Xuất file Word**, **Lưu dữ liệu (.json)**, **Mở dữ liệu (.json)**, **Xóa trắng** (có hỏi xác nhận).
- Tự lưu nháp vào localStorage (bọc try/catch), tự nạp lại khi mở trang.
- Ô "nếu có" chỉ hiện khi chọn "Có"/"Không bình thường".
- Code có chú thích tiếng Việt, chia rõ: dữ liệu cố định → đọc/ghi form → tạo Word → nút bấm.
