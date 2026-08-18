# tool-convert-md

Tool Python để chuyển đổi file Word/Excel sang Markdown, có hỗ trợ tách và liên kết ảnh.

## Tính năng

- Chuyển file `.docx` thành một file `.md` cùng tên.
- Chuyển file Excel (`.xlsx`, `.xlsm`) thành một thư mục cùng tên file.
- Quét toàn bộ sheet trong Excel, mỗi sheet xuất ra một file `.md`.
- Tự động trích xuất ảnh và tạo liên kết ảnh trong Markdown.
- Giữ cấu trúc output rõ ràng, dễ quản lý.

## Quy tắc output

### 1) Với DOCX

Input:

- `BaoCao.docx`

Output:

- `BaoCao.md`
- `BaoCao_images/` (chỉ tạo khi có ảnh)

### 2) Với Excel

Input:

- `DuLieu.xlsx`

Output:

- `DuLieu/`
- `DuLieu/<SheetName>.md` (mỗi sheet một file)
- `DuLieu/images/` (chỉ tạo khi có ảnh)

## Yêu cầu môi trường

- Python 3.9+

## Cài đặt

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Cách dùng

### Chuyển một file DOCX

```bash
python3 convert_to_md.py /duong-dan/toi/file.docx
```

### Chuyển một file Excel

```bash
python3 convert_to_md.py /duong-dan/toi/file.xlsx
```

### Chuyển toàn bộ file hỗ trợ trong một thư mục

```bash
python3 convert_to_md.py /duong-dan/toi/folder
```

### Chỉ định thư mục output

```bash
python3 convert_to_md.py /duong-dan/toi/file.xlsx -o /duong-dan/output
```

## Ghi chú

- Tool sẽ bỏ qua file tạm của Office có tiền tố `~$`.
- Với Excel, dữ liệu được xuất dạng bảng Markdown dựa trên vùng dữ liệu có nội dung.
- Ảnh trong Excel được gom vào thư mục `images` và chèn link tương đối trong file Markdown của sheet.

## File chính

- `convert_to_md.py`: script chuyển đổi
- `requirements.txt`: danh sách thư viện cần cài
