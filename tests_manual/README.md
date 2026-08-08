# Thư mục Test & Hướng dẫn Lấy Mail Outlook (`python-o365`)

Thư mục này chứa kịch bản kiểm thử việc lấy email từ Outlook và thao tác tạo thư mục mới thông qua thư viện `python-o365`.

## 1. Cấu trúc Kịch bản Test (`test_fetch_mail.py`)

Kịch bản [test_fetch_mail.py](file:///home/tien/code/python-o365/tests_manual/test_fetch_mail.py) thực hiện các bước:
1. **Khởi tạo Account**: Khởi tạo đối tượng `Account` với credentials và MSGraphProtocol.
2. **Lấy MailBox & Inbox**: Gọi `account.mailbox()` và `mailbox.inbox_folder()`.
3. **Lấy Email & Lọc Email (`get_messages`)**:
   - Sử dụng `query = mailbox.new_query()` để lọc mail chưa đọc (`isRead == False`).
   - Duyệt danh sách email, in tiêu đề, người gửi, xem trước nội dung và kiểm tra đính kèm.
4. **Tạo Thư mục mới**: Gọi `inbox.create_child_folder('Test_Mail_Folder')` để tạo folder mới trong hòm thư.

---

## 2. Cách Chạy Test

### 2.1. Chạy ở chế độ Giả lập (Mock Mode - Không cần API Key thật)
```bash
python3 tests_manual/test_fetch_mail.py
```

### 2.2. Chạy với Tài khoản Outlook / Azure Application thật
Chỉnh sửa file `test_fetch_mail.py` hoặc gọi hàm với credentials thật:
```python
from tests_manual.test_fetch_mail import test_outlook_mail_operations

test_outlook_mail_operations(
    client_id="YOUR_AZURE_CLIENT_ID",
    client_secret="YOUR_AZURE_CLIENT_SECRET",
    use_mock=False
)
```
*Lưu ý: Lần chạy đầu tiên sẽ yêu cầu bạn mở liên kết xác thực OAuth2 trong trình duyệt để đăng nhập tài khoản Microsoft.*
