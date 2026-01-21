# Metadata Tool (Python)

Ứng dụng desktop Tkinter hỗ trợ hai quy trình:
- **Combine**: gộp dữ liệu metadata_* từ thư mục nguồn vào file template để tạo bộ excel kết quả theo từng run/date.
- **Convert**: xuất dữ liệu từ file/folder Excel sang CSV (Sheet SampleImport và tùy chọn Aviti Manifest).

## Yêu cầu
- Python 3.10+ (đã có Tkinter).
- Thư viện: `openpyxl`.

Cài đặt nhanh:
```bash
pip install openpyxl
```

## Cấu trúc chính
- Backend/Funtion_Combie_Data.py — logic combine và ghi template.
- Fontend/combie.py — UI Combine (Tkinter).
- Fontend/convert.py — UI Convert + hàm `convert_path` xuất CSV.
- Fontend/main.py — điểm vào ứng dụng, đăng ký các màn hình.
- Fontend/settings.json — lưu đường dẫn gần nhất người dùng chọn.

## Chạy ứng dụng
```bash
cd "c:/Users/VUDUCHAI/OneDrive - Gene Solutions/Tool Excel Team Bi/Metadata Tool_ Python/Fontend"
python main.py
```
Ứng dụng mở cửa sổ Tkinter với hai lựa chọn Combine và Convert.

## Hướng dẫn nhanh
### Combine
1) Chọn **Source Folder** chứa các file `metadata_*.xlsx`.
2) Chọn **Template File** (file Excel mẫu).
3) Chọn **Output Folder** nơi lưu kết quả.
4) Nhấn **Combine Data**. Xong sẽ hiện nút mở thư mục output. Đồng thời, đường dẫn output được gán cho màn hình Convert làm input folder.

### Convert
- Chế độ file: chọn **Input File** (xlsx) và **Output Folder**, nhấn **Convert File**.
- Chế độ folder: chọn **Input Folder** (tự động quét *.xlsx) và **Output Folder**, nhấn **Convert Folder**.
- Tùy chọn **Combie duo** bật thêm xuất sheet "Aviti Manifest" (bắt đầu hàng 16) ngoài sheet chính (SampleImport, bắt đầu hàng 24).
- Tiến độ hiển thị theo %; sau khi xong có nút mở thư mục output.

## Lưu ý sử dụng
- Bỏ qua file Excel tạm của Office (`~$`).
- Thư mục output sẽ được tạo nếu chưa tồn tại.
- Nếu thiếu sheet yêu cầu, ứng dụng sẽ báo lỗi.

## Gỡ lỗi nhanh
- Đảm bảo đã cài `openpyxl` và chạy bằng đúng bản Python có Tkinter.
- Nếu giao diện không mở: kiểm tra lệnh chạy ở đúng thư mục `Fontend`.
- Nếu không thấy file CSV/Excel mới: kiểm tra đường dẫn output và quyền ghi.

## Giấy phép
Dùng nội bộ (chưa khai báo license).
