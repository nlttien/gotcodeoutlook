"""
Kịch bản Test: Lấy Email từ Outlook và Thao tác Thư mục Mail với O365 Python Library.
"""

import sys
from unittest.mock import MagicMock

try:
    from O365 import Account, MSGraphProtocol
except ModuleNotFoundError as err:
    print(f"[!] Thiếu thư viện phụ thuộc ({err}). Vui lòng cài đặt bằng lệnh:")
    print("    pip install O365 msal requests_oauthlib beautifulsoup4 python-dateutil")
    sys.exit(0)


def test_outlook_mail_operations(client_id="MOCK_CLIENT_ID", client_secret="MOCK_CLIENT_SECRET", use_mock=True):
    print("=" * 60)
    print(" 1. KHỞI TẠO TÀI KHOẢN (ACCOUNT)")
    print("=" * 60)

    credentials = (client_id, client_secret)
    protocol = MSGraphProtocol(api_version='v1.0')
    account = Account(credentials, protocol=protocol)

    if use_mock:
        print("[MOCK] Đang giả lập kết nối Microsoft Graph Connection...")
        account.con = MagicMock()
        # Giả lập trạng thái đã xác thực
        account.con.token_backend.has_data = True
        account.con.token_backend.token_is_expired.return_value = False
        account.con.token_backend.token_is_long_lived.return_value = True

    print(f"Xác thực tài khoản: {'Thành công' if account.is_authenticated else 'Chưa xác thực'}")

    print("\n" + "=" * 60)
    print(" 2. TRUY CẬP HÒM THƯ (MAILBOX & FOLDERS)")
    print("=" * 60)

    mailbox = account.mailbox()
    print(f"Khởi tạo MailBox resource: {mailbox.main_resource}")

    # Lấy thư mục Inbox
    inbox = mailbox.inbox_folder()
    print(f"Thư mục Inbox: {inbox.name if hasattr(inbox, 'name') else 'Inbox'}")

    print("\n" + "=" * 60)
    print(" 3. ĐỌC / LẤY DANH SÁCH EMAIL (GET MESSAGES)")
    print("=" * 60)

    if use_mock:
        # Giả lập response Graph API trả về danh sách email
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'value': [
                {
                    'id': 'msg_001',
                    'subject': 'Thông báo họp khẩn cấp',
                    'bodyPreview': 'Nội dung cuộc họp tuần này...',
                    'sender': {'emailAddress': {'name': 'Nguyen Van A', 'address': 'a.nguyen@example.com'}},
                    'isRead': False,
                    'hasAttachments': True,
                    'createdDateTime': '2026-08-08T09:00:00Z'
                },
                {
                    'id': 'msg_002',
                    'subject': 'Báo cáo doanh thu tháng 7',
                    'bodyPreview': 'Chi tiết báo cáo đính kèm bên dưới...',
                    'sender': {'emailAddress': {'name': 'Tran Thi B', 'address': 'b.tran@example.com'}},
                    'isRead': True,
                    'hasAttachments': False,
                    'createdDateTime': '2026-08-07T14:30:00Z'
                }
            ]
        }
        account.con.get.return_value = mock_response

    # Tạo truy vấn lọc email bằng QueryBuilder (v2.1+)
    q = mailbox.q()
    query = q.equals('isRead', False)  # Lấy mail chưa đọc
    
    print("Đang tải danh sách tin nhắn từ Inbox...")
    messages = list(inbox.get_messages(limit=5, query=query))
    print(f"Số lượng email nhận được: {len(messages)}")

    for idx, msg in enumerate(messages, 1):
        print(f"\n--- Email #{idx} ---")
        print(f"  Tiêu đề (Subject) : {msg.subject}")
        print(f"  Người gửi (Sender): {msg.sender}")
        print(f"  Có tệp đính kèm    : {msg.has_attachments}")
        print(f"  Nội dung xem trước : {msg.body_preview if hasattr(msg, 'body_preview') else msg.body}")

    print("\n" + "=" * 60)
    print(" 4. TẠO THƯ MỤC MỚI (CREATE FOLDER)")
    print("=" * 60)

    new_folder_name = "Test_Mail_Folder"
    print(f"Đang tạo thư mục mới '{new_folder_name}' trong Inbox...")

    if use_mock:
        mock_folder_res = MagicMock()
        mock_folder_res.status_code = 201
        mock_folder_res.json.return_value = {
            'id': 'folder_999',
            'displayName': new_folder_name,
            'childFolderCount': 0,
            'unreadItemCount': 0,
            'totalItemCount': 0
        }
        account.con.post.return_value = mock_folder_res

    try:
        new_folder = inbox.create_child_folder(new_folder_name)
        print(f"Tạo thư mục thành công! Tên: {new_folder.name}, ID: {new_folder.folder_id}")
    except Exception as e:
        print(f"Lỗi khi tạo thư mục: {e}")

    print("\n" + "=" * 60)
    print(" HOÀN THÀNH TEST LẤY MAIL VÀ TẠO THƯ MỤC OUTLOOK")
    print("=" * 60)


if __name__ == '__main__':
    # Chạy ở chế độ Mock mặc định để xác minh code không lỗi
    test_outlook_mail_operations(use_mock=True)
