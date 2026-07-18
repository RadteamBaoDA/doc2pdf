# doc2pdf

[English](../README.md)

CLI cho Windows dùng Microsoft Office để chuyển đổi Word, Excel và PowerPoint sang PDF. Ứng dụng hỗ trợ file đơn, thư mục đệ quy, cấu hình theo mẫu đường dẫn, báo cáo lỗi và hậu xử lý PDF.

Ngoài chế độ PDF, command `convert-macros` chỉ dùng để tạo bản không macro:

- `.docm` → `.docx`
- `.pptm` → `.pptx`
- `.xlsm` → `.xlsx`

> Việc chuyển sang định dạng không macro sẽ loại bỏ VBA project. File nguồn không bị sửa.

## Yêu cầu

- Windows
- Python 3.12+
- Microsoft Word, Excel và/hoặc PowerPoint đã được cài đặt tương ứng với loại file cần xử lý

## Cài đặt

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install .
```

Chế độ phát triển:

```powershell
python -m pip install -e ".[dev]"
```

## Tất cả command và mode chạy

### Xem trợ giúp và phiên bản

```powershell
doc2pdf --help
doc2pdf --version
doc2pdf convert --help
doc2pdf convert-macros --help
```

Có thể chạy trực tiếp từ source mà không cần cài command:

```powershell
python -m src.cli --help
```

### Mode 1: Office → PDF

Chuyển một file, output mặc định là thư mục `output`:

```powershell
doc2pdf convert "input\report.docx"
```

Chỉ định file PDF đầu ra:

```powershell
doc2pdf convert "input\report.docx" --output "output\report.pdf"
```

Chuyển toàn bộ thư mục đệ quy và giữ cấu trúc thư mục con:

```powershell
doc2pdf convert "input" --output "output"
```

Dùng file cấu hình khác:

```powershell
doc2pdf convert "input" --output "output" --config "config.yml"
```

Bật log chi tiết, bật/tắt trim khoảng trắng hoặc ghi đè margin trim:

```powershell
doc2pdf convert "input" --output "output" --verbose
doc2pdf convert "input" --output "output" --trim
doc2pdf convert "input" --output "output" --no-trim
doc2pdf convert "input" --output "output" --trim --trim-margin 10
```

Các input được nhận diện: `.doc`, `.docx`, `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.ppt`, `.pptx` và `.pdf`. Cách xử lý PDF có sẵn trong input được điều khiển bởi `pdf_handling` trong `config.yml`.

### Mode 2: Loại bỏ macro, không tạo PDF

Một file (output mặc định trong thư mục `output`):

```powershell
doc2pdf convert-macros "input\report.docm"
doc2pdf convert-macros "input\slides.pptm"
doc2pdf convert-macros "input\workbook.xlsm"
```

Chỉ định chính xác file đầu ra:

```powershell
doc2pdf convert-macros "input\report.docm" --output "clean\report.docx"
```

Chuyển hỗn hợp cả ba loại trong toàn bộ thư mục, giữ cấu trúc thư mục con:

```powershell
doc2pdf convert-macros "input" --output "clean"
```

Command này chỉ đọc `.docm`, `.pptm`, `.xlsm`; các file khác trong thư mục được bỏ qua. Macro bị vô hiệu hóa khi Office mở file bằng automation và không được lưu vào file đầu ra.

## Cấu hình

`config.yml` là cấu hình mặc định. Các nhóm chính hiện có gồm:

- `timeout`: timeout chuyển đổi tài liệu và trim Excel.
- `logging`: mức log, file log, rotation và retention.
- `post_processing`: trim khoảng trắng sau khi tạo PDF.
- `suffix`: hậu tố tên PDF theo Word, Excel, PowerPoint.
- `reporting`: báo cáo kết quả, lỗi và sao chép file lỗi.
- `pdf_handling`: cách xử lý PDF đã có trong input.
- các thiết lập PDF/layout riêng cho Word, Excel và PowerPoint, bao gồm rule theo pattern.

`convert-macros` không dùng các thiết lập PDF trong `config.yml`.

## Kiểm thử

```powershell
python -m pytest
```

## Lưu ý vận hành

- Đóng tài liệu Office đang mở trước khi chạy batch để tránh file lock hoặc hộp thoại COM.
- Không mở file đầu ra trùng lúc chuyển đổi.
- Log, summary và error report được ghi theo `config.yml`.
- Nếu PowerShell không nhận `doc2pdf`, hãy kích hoạt virtual environment hoặc dùng `python -m src.cli` thay cho `doc2pdf` trong các ví dụ trên.
