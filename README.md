# TRỢ LÝ TẠO ĐỀ KIỂM TRA THEO CÔNG VĂN 7991/BGDĐT-GDTrH
### GVBM: Lê Tâm - Trường THCS Quang Trung

Ứng dụng AI thông minh hỗ trợ giáo viên xây dựng bộ hồ sơ kiểm tra định kỳ hoàn chỉnh theo chuẩn **Công văn 7991/BGDĐT-GDTrH** (ngày 17/12/2024 của Bộ Giáo dục và Đào tạo).

---

## 🌟 Tính Năng Nổi Bật

1. **Tài liệu Sách giáo khoa (SGK) chuẩn**: Cho phép tải lên các chương, bài học từ file SGK (PDF, Word, TXT). AI sẽ bám sát 100% kiến thức, định nghĩa, số liệu, ngữ liệu và ví dụ minh họa chính xác từ SGK.
2. **Khung Kế hoạch bài dạy (Phần 1)**: Xây dựng mục tiêu (kiến thức, năng lực, phẩm chất), thiết bị dạy học, tiến trình bài dạy bám sát nội dung Phụ lục 3.
3. **Ma trận đề kiểm tra (Phần 2)**: Chuẩn bảng HTML định dạng rowspan/colspan phân bổ chuẩn 3 mức độ nhận thức (Biết - Hiểu - Vận dụng) với tổng tỉ lệ khớp tuyệt đối 100%.
4. **Bản đặc tả đề kiểm tra (Phần 3)**: Tách bạch rõ ràng Yêu cầu cần đạt theo từng mức độ (Biết - Hiểu - Vận dụng) và vị trí câu hỏi tương ứng.
5. **Bộ đề kiểm tra chuẩn format mới (Phần 4 & Phần 5)**:
   - **Phần I**: Trắc nghiệm nhiều phương án lựa chọn (4 lựa chọn A, B, C, D trên từng dòng).
   - **Phần II**: Trắc nghiệm Đúng - Sai (mỗi câu 4 lệnh hỏi a, b, c, d; tính điểm chuẩn Bộ GDĐT).
   - **Phần III**: Trắc nghiệm trả lời ngắn (kết quả ngắn gọn, số hoặc đơn vị).
   - **Phần IV**: Tự luận (kèm bảng ma trận barem chấm chi tiết từng bước).
   - Đề số 2 tương đương hoàn toàn về cấu trúc và độ khó.
6. **Quản lý nhiều API Key & Tự động xoay vòng (Failover Engine)**: Hỗ trợ dán hàng loạt API Key, tự động chuyển key khi gặp lỗi quá tải hạn ngạch (Rate Limit 429).
7. **Hỗ trợ công thức Toán học & Hình học SVG**:
   - Trình bày công thức chuẩn LaTeX / KaTeX / MathType.
   - Minh họa hình vẽ trực quan bằng mã SVG chuẩn xác.
8. **Xuất file nhanh chóng**:
   - Xuất toàn bộ hồ sơ ra file Word (.docx) chuẩn font Times New Roman, hỗ trợ MathType.
   - Xuất bảng ma trận và đặc tả sang file Excel / Trang tính (.xlsx).

---

## 🚀 Hướng Dẫn Cài Đặt & Chạy Cục Bộ

### Yêu cầu:
- **Node.js** (phiên bản 18 trở lên).
- **Google Gemini API Key** (lấy miễn phí tại [Google AI Studio](https://aistudio.google.com/app/apikey)).

### Các bước thực hiện:

1. **Cài đặt thư viện phụ thuộc:**
   ```bash
   npm install
   ```

2. **Cấu hình API Key:**
   - Tạo file `.env.local` hoặc nhập trực tiếp một hoặc nhiều API Key trên giao diện web:
     ```env
     GEMINI_API_KEY="AIzaSy..."
     ```

3. **Khởi chạy ứng dụng:**
   ```bash
   npm run dev
   ```
   Mở trình duyệt truy cập: [http://localhost:3005](http://localhost:3005)

---
*Phát triển phục vụ công tác chuyên môn giảng dạy tại Trường THCS Quang Trung - GVBM: Lê Tâm.*
