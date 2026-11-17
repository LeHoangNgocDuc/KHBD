# Smart KHBD Toán 9

Ứng dụng mẫu chạy cục bộ để nạp PPCT, SGK PDF và KHBD Word; xem dữ liệu trích xuất; bật/tắt NLS, AI, GDHN; xuất hoặc ghép giáo án DOCX.

## Cài đặt và chạy

```bash
cd lesson-planner
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt
uvicorn backend.main:app --reload
```

Mở `http://127.0.0.1:8000`. Dữ liệu tải lên được lưu trong thư mục tạm `backend/data` của máy chạy app. Đây là bản local, chưa có đăng nhập, phân quyền hay lưu trữ nhiều người dùng.

## Cấu hình nhận diện FastAPI trên Vercel

Tệp `pyproject.toml` đã khai báo `backend.main:app` làm FastAPI entrypoint. Khi import repository lên Vercel, đặt **Root Directory** tại thư mục có `pyproject.toml`, `requirements.txt` và thư mục `backend/`.

Lưu ý: khai báo entrypoint chỉ xử lý lỗi build `FASTAPI_ENTRYPOINT_NOT_FOUND`. Bản này vẫn cần thay đổi cách lưu tệp/phiên và cách nhận PDF lớn trước khi dùng đủ tính năng trên Vercel Functions.

## Chức năng hiện có

- Nạp PDF/DOCX/XLSX PPCT; parser lấy tuần, khoảng ngày, tiết PPCT, tên bài và phân môn theo nội dung văn bản. Có thể sửa trường hợp bảng PDF bị tách dòng sai.
- Nạp SGK PDF; trích xuất văn bản theo trang, danh sách ảnh nhúng và mục lục heuristic. Ảnh có thể tải riêng; ghi chú fallback gồm trang SGK.
- Dùng DOCX tải lên làm mẫu cơ sở, cập nhật ngày soạn, tuần, lớp, tên bài; bật/tắt khối NLS/AI và GDHN; chèn dòng tham chiếu bài học kéo dài qua tuần.
- Có thể tự nhập JSON cấu trúc bài dạy; hình SGK được gắn vào nội dung qua anchor xác định. Cấu hình ảnh trong `lesson_content.images`: `[ {"page": 10, "description": "Đồ thị hệ phương trình", "anchor": "Ta có hình biểu diễn nghiệm"} ]`.
- Sinh DOCX mẫu 5512 có bảng hoạt động 2 cột từ dữ liệu có cấu trúc; ghép DOCX bằng `docxcompose`.
- Chuyển LaTeX inline cơ bản sang OMML editable. Converter không raster hóa công thức; cú pháp chưa hỗ trợ sẽ báo lỗi để giáo viên sửa.

## Dữ liệu đi kèm đã đối chiếu

Ứng dụng được thiết kế để dùng với bộ tải lên gồm `DS9_T1_T1_2_B1.docx`, `KHDH Toán 9(26-27)-chia theo tuần(1).pdf`, và hai tập SGK Toán 9. Có thể nạp lại bộ khác qua giao diện. Không đóng gói lại các sách/giáo án nguồn trong mã nguồn.

## Giới hạn cần biết

- “Độ chính xác 100%” về hình thức chỉ đạt được khi dùng DOCX chuẩn của chính giáo viên và giữ nội dung/bảng nguồn; bộ sinh cấu trúc chung tái tạo cấu trúc chứ không thể cam kết trùng tuyệt đối mọi style, ngắt trang, font nhúng và vị trí hình của một mẫu chưa được cấu hình.
- PPCT PDF có thể có ô gộp hoặc chữ nằm trên nhiều dòng; kết quả có thể cần hiệu chỉnh trong bảng xem trước.
- OCR không chạy với PDF scan. Mã hiện tại nhận dạng văn bản PDF; trang không có text vẫn trả về trang và fallback để người dùng bổ sung mô tả.
- Không gọi mô hình AI bên ngoài. Tính năng AI là các chỗ tích hợp/prompt trong giáo án, không phải tự sinh nội dung.
- XSLT chính thức MML2OMML có thể được đặt tại `backend/MML2OMML.XSL`; nếu không có, converter dùng bộ chuyển đổi MathML subset đã tích hợp.

## Cấu trúc

```text
lesson-planner/
├── backend/
│   ├── main.py
│   ├── docx_generator.py
│   ├── omml_converter.py
│   ├── pdf_parser.py
│   ├── ppct_parser.py
│   ├── merger.py
│   └── models.py
├── frontend/index.html
├── requirements.txt
└── README.md
```
