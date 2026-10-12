# App Streamlit tra cứu điểm tự đọc từ thư mục

## Cấu trúc thư mục
Giải nén, sau đó đặt file điểm như sau:

```text
tra_cuu_diem_tu_thu_muc/
├── home.py
├── requirements.txt
├── README.md
└── DuLieuDiem/
    ├── 6A1.xlsx
    ├── 6A2.xlsx
    └── 6A3.xlsx
```

Mỗi file Excel là một lớp; mỗi sheet là một môn học. App mặc định đọc thư mục `DuLieuDiem` nằm cạnh `home.py`. Có thể đổi thư mục trong thanh bên của app.

## Cài đặt và chạy trên Windows
Mở PowerShell tại thư mục có `home.py`, chạy:

```powershell
py -m pip install -r requirements.txt
py -m streamlit run home.py
```

Nếu máy không nhận lệnh `py`, thử thay `py` bằng `python`.

## Cập nhật điểm
Khi cập nhật các file Excel, lưu file vào thư mục `DuLieuDiem`, sau đó bấm **Đọc lại dữ liệu** trong thanh bên của app. Nếu app chưa phản ánh thay đổi, tải lại trang trình duyệt.

## Lưu ý quan trọng
- App đọc các file `.xlsx` trong thư mục trên **máy đang chạy Streamlit**. Nếu triển khai lên máy chủ, cần đặt file điểm trên máy chủ đó hoặc cấu hình thư mục dùng chung được máy chủ truy cập.
- Đóng Excel hoặc đảm bảo file đã lưu trước khi bấm đọc lại.
- Mã định danh và điểm là dữ liệu cá nhân học sinh. Chỉ cho phép truy cập app trong phạm vi được nhà trường phê duyệt; không mở công khai khi chưa có xác thực người dùng.
