"""
API Server & Web UI: Dịch vụ trích xuất mã xác nhận (OTP) và hiển thị hòm thư Outlook bằng python-o365 + FastAPI.
"""

import re
import sys
import sqlite3
from pathlib import Path
from typing import Optional, List, Union
from unittest.mock import MagicMock

try:
    from fastapi import FastAPI, HTTPException
    from fastapi.middleware.cors import CORSMiddleware
    from fastapi.staticfiles import StaticFiles
    from fastapi.responses import FileResponse
    from pydantic import BaseModel, Field
    import uvicorn
    from O365 import Account, MSGraphProtocol

except ModuleNotFoundError as err:
    print(f"[!] Thiếu thư viện phụ thuộc ({err}). Vui lòng chạy lệnh:")
    print("    pip install fastapi uvicorn O365 msal requests_oauthlib beautifulsoup4 python-dateutil")
    sys.exit(1)

DB_PATH = Path(__file__).parent / "accounts.db"


def init_sqlite_db():
    """Khởi tạo duy nhất 1 bảng accounts duy nhất trong SQLite Database"""
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS accounts (
                email TEXT PRIMARY KEY,
                account_str TEXT NOT NULL,
                status TEXT DEFAULT 'Hoạt động',
                otp_code TEXT,
                subject TEXT,
                sender TEXT,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        # Đảm bảo bổ sung các cột nếu bảng cũ chưa có
        cursor.execute("PRAGMA table_info(accounts)")
        columns = [col[1] for col in cursor.fetchall()]
        if "status" not in columns:
            cursor.execute("ALTER TABLE accounts ADD COLUMN status TEXT DEFAULT 'Hoạt động'")
        if "otp_code" not in columns:
            cursor.execute("ALTER TABLE accounts ADD COLUMN otp_code TEXT")
        if "subject" not in columns:
            cursor.execute("ALTER TABLE accounts ADD COLUMN subject TEXT")
        if "sender" not in columns:
            cursor.execute("ALTER TABLE accounts ADD COLUMN sender TEXT")
        conn.commit()


def save_account_to_db(account_str: str, status: Optional[str] = None, otp_code: Optional[str] = None, subject: Optional[str] = None, sender: Optional[str] = None) -> Optional[str]:
    """Lưu / Cập nhật thông tin tài khoản, danh sách trạng thái và mã OTP mới nhất vào duy nhất 1 bảng accounts trong SQLite DB"""
    acc = parse_account_string(account_str)
    email = acc['username'].strip().lower()
    if not email or '@' not in email:
        return None
    
    # Xử lý chuỗi status (nếu dạng list hoặc tuple thì join)
    if isinstance(status, (list, tuple)):
        status_str = ", ".join([str(s).strip() for s in status if str(s).strip()])
    elif status:
        status_str = str(status).strip()
    else:
        status_str = None

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        
        # Nếu chuỗi mới truyền vào không chứa dấu |, kiểm tra xem SQLite DB đã có chuỗi token đầy đủ chưa
        if '|' not in account_str:
            cursor.execute("SELECT account_str FROM accounts WHERE LOWER(email) = LOWER(?)", (email,))
            existing = cursor.fetchone()
            if existing and '|' in existing[0]:
                account_str = existing[0]

        cursor.execute("""
            INSERT INTO accounts (email, account_str, status, otp_code, subject, sender, updated_at)
            VALUES (?, ?, COALESCE(?, 'Hoạt động'), ?, ?, ?, datetime('now'))
            ON CONFLICT(email) DO UPDATE SET
                account_str = excluded.account_str,
                status = COALESCE(excluded.status, accounts.status, 'Hoạt động'),
                otp_code = COALESCE(excluded.otp_code, accounts.otp_code),
                subject = COALESCE(excluded.subject, accounts.subject),
                sender = COALESCE(excluded.sender, accounts.sender),
                updated_at = datetime('now')
        """, (email, account_str, status_str, otp_code, subject, sender))
        conn.commit()
    return email


def get_account_from_db(query_key: str) -> Optional[str]:
    """Tra cứu chuỗi tài khoản từ SQLite DB (Case-insensitive multi-field search per User Rule #4)"""
    clean_key = query_key.strip().lower()
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT account_str FROM accounts 
            WHERE LOWER(email) = LOWER(?) OR LOWER(account_str) LIKE LOWER(?)
            ORDER BY updated_at DESC LIMIT 1
        """, (clean_key, f"%{clean_key}%"))
        row = cursor.fetchone()
        if row:
            return row[0]
    return Nonelean_key, f"%{clean_key}%"))
        row = cursor.fetchone()
        if row:
            return row[0]
    return None


init_sqlite_db()

app = FastAPI(
    title="Outlook Mail & OTP Extractor Dashboard",
    description="API Server và Giao diện Web lấy mã OTP 1-click & xem toàn bộ hòm thư Outlook.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Phục vụ các tệp tĩnh (CSS, JS, Images)
static_dir = Path(__file__).parent / "static"
static_dir.mkdir(parents=True, exist_ok=True)
app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")



class AccountCodeRequest(BaseModel):
    account_str: str = Field(
        ...,
        description="Dòng thông tin tài khoản dạng: email|password|refresh_token|client_id"
    )
    keyword: Optional[str] = Field(
        default=None,
        description="Từ khóa lọc email (VD: OTP, code, mã xác nhận, Facebook...)"
    )
    limit: int = Field(default=15, ge=1, le=50, description="Số lượng email gần nhất cần quét")
    use_mock: bool = Field(default=False, description="Đặt True để chạy test dữ liệu giả lập, False để kết nối mail thật")


class EmailItem(BaseModel):
    id: str
    subject: str
    sender: str
    created_date: str
    body_preview: str
    body: str = ""
    has_attachments: bool
    otp_codes: List[str] = []



class OTPCodeResponse(BaseModel):
    status: str
    email: str
    otp_code: Optional[str] = None
    subject: Optional[str] = None
    sender: Optional[str] = None
    received_time: Optional[str] = None
    message_preview: Optional[str] = None
    all_codes_found: List[str] = []
    all_messages: List[EmailItem] = []


def parse_account_string(account_str: str) -> dict:
    """Bóc tách dòng thông tin tài khoản định dạng: email|password|refresh_token|client_id"""
    parts = [p.strip() for p in account_str.split('|')]
    rf = parts[2] if len(parts) > 2 else ''
    # Loại bỏ các dấu $ thừa ở cuối chuỗi token (nếu có)
    rf = rf.rstrip('$')
    return {
        'username': parts[0] if len(parts) > 0 else '',
        'password': parts[1] if len(parts) > 1 else '',
        'refresh_token': rf,
        'client_id': parts[3] if len(parts) > 3 else ''
    }



def extract_otp_code(text: str) -> List[str]:
    """Trích xuất danh sách mã xác nhận (4-8 chữ số hoặc mã dạng gạch ngang ngắn) từ văn bản"""
    if not text:
        return []
    
    # 1. Ưu tiên hàng đầu cho các mã số OTP chuẩn (4-8 chữ số), loại bỏ các năm 2024-2027
    cleaned_text = re.sub(r'\b20[23]\d\b', '', text)
    digit_codes = re.findall(r'\b\d{4,8}\b', cleaned_text)

    # 2. Tìm các mã dạng gạch ngang ngắn (VD: ABC-123, 2F4-07E) - loại bỏ UUIDs và Client IDs dài
    hyphen_matches = re.findall(r'\b[a-zA-Z0-9]{2,6}(?:-[a-zA-Z0-9]{2,6})+\b', text)
    hyphen_codes = [c for c in hyphen_matches if len(c) <= 16 and not c.startswith('9e5f')]

    # Mã số đơn thuần được sắp trước, sau đó tới mã gạch ngang
    all_matches = digit_codes + hyphen_codes

    seen = set()
    unique_codes = []
    for code in all_matches:
        if code not in seen:
            seen.add(code)
            unique_codes.append(code)
    return unique_codes




def delete_account_from_db(email: str) -> bool:
    """Xóa tài khoản khỏi SQLite Database theo email (Case-insensitive)"""
    clean_email = email.strip().lower()
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM accounts WHERE LOWER(email) = LOWER(?)", (clean_email,))
        conn.commit()
        return cursor.rowcount > 0


class AddAccountRequest(BaseModel):
    account_str: str = Field(..., description="Dòng thông tin tài khoản dạng email|password|refresh_token|client_id")


@app.get("/")
def read_root():
    """Phục vụ trang Giao diện Web UI Dashboard"""
    index_file = static_dir / "index.html"
    if index_file.exists():
        return FileResponse(index_file)
    return {"message": "Web UI chưa khởi tạo. Vui lòng tạo static/index.html"}


@app.get("/admin")
def read_admin():
    """Phục vụ trang Giao diện Quản trị SQLite Database Admin"""
    admin_file = static_dir / "admin.html"
    if admin_file.exists():
        return FileResponse(admin_file)
    return {"message": "Admin Web UI chưa khởi tạo. Vui lòng tạo static/admin.html"}


@app.get("/api/status")
def get_server_status():
    """Endpoint kiểm tra trạng thái hoạt động của Python O365 API Server per User Rule #5"""
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM accounts")
        total_accounts = cursor.fetchone()[0]
    return {
        "status": "success",
        "service": "Python O365 API Server",
        "total_saved_accounts": total_accounts,
        "db_path": str(DB_PATH)
    }


@app.get("/api/accounts")
def list_saved_accounts(search: Optional[str] = None, status: Optional[str] = None):
    """Lấy danh sách tất cả tài khoản, lọc tìm kiếm đa trường và danh sách trạng thái từ duy nhất 1 bảng accounts trong SQLite DB"""
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT email, account_str, status, otp_code, subject, sender, updated_at FROM accounts ORDER BY updated_at DESC")
        rows = cursor.fetchall()
        
        items = []
        for r in rows:
            acc_status = r[2] or "Hoạt động"
            item = {
                "email": r[0],
                "account_str": r[1],
                "status": acc_status,
                "otp_code": r[3],
                "subject": r[4],
                "sender": r[5],
                "updated_at": r[6]
            }
            
            # Lọc tìm kiếm search (Case-insensitive multi-field search per User Rule #4)
            if search:
                s_lower = search.strip().lower()
                matched = (
                    s_lower in (r[0] or "").lower() or
                    s_lower in (r[1] or "").lower() or
                    s_lower in (acc_status or "").lower() or
                    s_lower in (r[3] or "").lower() or
                    s_lower in (r[4] or "").lower() or
                    s_lower in (r[5] or "").lower()
                )
                if not matched:
                    continue

            # Lọc theo status (Hỗ trợ lọc chuỗi chứa từ khóa status per User Rule #3)
            if status:
                st_lower = status.strip().lower()
                if st_lower not in acc_status.lower():
                    continue

            items.append(item)

        return {"status": "success", "accounts": items, "count": len(items)}


class AddAccountRequest(BaseModel):
    account_str: str = Field(..., description="Dòng thông tin tài khoản dạng email|password|refresh_token|client_id")
    status: Optional[Union[str, List[str]]] = Field(default=None, description="Trạng thái hoặc danh sách trạng thái của tài khoản")


class UpdateAccountRequest(BaseModel):
    email: str = Field(..., description="Địa chỉ email tài khoản cần cập nhật")
    account_str: Optional[str] = Field(None, description="Dòng thông tin tài khoản mới")
    status: Optional[Union[str, List[str]]] = Field(None, description="Danh sách trạng thái hoặc chuỗi trạng thái mới")


class BatchAddAccountsRequest(BaseModel):
    accounts: List[str] = Field(..., description="Danh sách chuỗi tài khoản dạng mảng hoặc văn bản nhiều dòng")
    status: Optional[Union[str, List[str]]] = Field(default=None, description="Trạng thái áp dụng chung cho batch")


@app.post("/api/accounts/add")
def add_new_account(req: AddAccountRequest):
    """Thêm hoặc cập nhật tài khoản vào SQLite Database (Hỗ trợ 1 hoặc nhiều dòng)"""
    raw_str = req.account_str.strip()
    if not raw_str:
        raise HTTPException(status_code=400, detail="Vui lòng nhập chuỗi tài khoản hợp lệ.")

    lines = [l.strip() for l in raw_str.splitlines() if l.strip()]
    if len(lines) == 1:
        email = save_account_to_db(account_str=lines[0], status=req.status)
        if not email:
            raise HTTPException(status_code=400, detail="Chuỗi account_str không chứa địa chỉ email hợp lệ.")
        return {"status": "success", "message": f"Đã lưu tài khoản {email} vào SQLite Database", "email": email}
    else:
        return batch_add_accounts(BatchAddAccountsRequest(accounts=lines, status=req.status))


@app.put("/api/accounts/update")
@app.put("/api/accounts/{email:path}")
def update_account(req: UpdateAccountRequest, email: Optional[str] = None):
    """Cập nhật thông tin chuỗi script key và danh sách trạng thái của email trong SQLite DB"""
    target_email = email or req.email
    if not target_email or '@' not in target_email:
        raise HTTPException(status_code=400, detail="Vui lòng cung cấp địa chỉ email hợp lệ.")

    # Tìm account_str cũ nếu người dùng không truyền account_str mới
    existing_acc_str = get_account_from_db(target_email)
    account_str_to_save = req.account_str or existing_acc_str or target_email

    saved_email = save_account_to_db(account_str=account_str_to_save, status=req.status)
    if not saved_email:
        raise HTTPException(status_code=400, detail=f"Không thể cập nhật tài khoản {target_email}.")

    return {
        "status": "success",
        "message": f"Đã cập nhật thành công tài khoản {saved_email}",
        "email": saved_email
    }


@app.post("/api/accounts/batch")
def batch_add_accounts(req: BatchAddAccountsRequest):
    """Import hàng loạt chuỗi tài khoản Outlook vào SQLite Database"""
    added_count = 0
    failed_count = 0
    processed_emails = []

    raw_lines = []
    for item in req.accounts:
        if isinstance(item, str):
            for line in item.splitlines():
                line = line.strip()
                if line:
                    raw_lines.append(line)

    for line in raw_lines:
        try:
            email = save_account_to_db(account_str=line, status=req.status)
            if email:
                added_count += 1
                processed_emails.append(email)
            else:
                failed_count += 1
        except Exception as e:
            print(f"Lỗi import dòng '{line[:30]}...': {e}")
            failed_count += 1

    return {
        "status": "success",
        "message": f"Đã import thành công {added_count} tài khoản vào SQLite Database" + (f" ({failed_count} dòng lỗi/không hợp lệ)" if failed_count > 0 else ""),
        "added_count": added_count,
        "failed_count": failed_count,
        "emails": processed_emails
    }


class EmailOnlyRequest(BaseModel):
    email: str = Field(..., description="Địa chỉ email tài khoản Outlook (VD: sylvesterrojas997795@outlook.com)")
    keyword: Optional[str] = Field(None, description="Từ khóa lọc email")
    limit: Optional[int] = Field(15, description="Số lượng email tối đa cần đọc")


@app.delete("/api/accounts/delete/{email:path}")
@app.delete("/api/accounts/{email:path}")
def delete_account(email: str):
    """Xóa tài khoản khỏi SQLite Database theo email"""
    if email in ("batch", "add"):
        raise HTTPException(status_code=400, detail="Tên email không hợp lệ.")
    success = delete_account_from_db(email)
    if not success:
        raise HTTPException(status_code=404, detail=f"Không tìm thấy tài khoản '{email}' trong SQLite Database.")
    return {"status": "success", "message": f"Đã xóa tài khoản {email} khỏi SQLite Database"}



@app.get("/api/get-code-by-email")
def get_code_by_email_get(email: str, keyword: Optional[str] = None, limit: int = 15):
    """
    Endpoint mới (GET): Chỉ cần truyền query param ?email=your_email@outlook.com để nhận mã OTP.
    Ví dụ: GET /api/get-code-by-email?email=sylvesterrojas997795@outlook.com
    """
    req = AccountCodeRequest(account_str=email, keyword=keyword, limit=limit, use_mock=False)
    return get_verification_code(req)


@app.post("/api/get-code-by-email")
def get_code_by_email_post(req: EmailOnlyRequest):
    """
    Endpoint mới (POST): Chỉ cần truyền body JSON {"email": "your_email@outlook.com"} để nhận mã OTP.
    Ví dụ: POST /api/get-code-by-email với body {"email": "sylvesterrojas997795@outlook.com"}
    """
    code_req = AccountCodeRequest(account_str=req.email, keyword=req.keyword, limit=req.limit or 15, use_mock=False)
    return get_verification_code(code_req)


@app.post("/api/get-code", response_model=OTPCodeResponse)
def get_verification_code(req: AccountCodeRequest):
    """
    Nhận chuỗi tài khoản email|password|refresh_token|client_id (hoặc tên email để tra cứu từ SQLite DB),
    đọc hòm thư Outlook và trả về mã OTP cùng toàn bộ danh sách email.
    """

    raw_input = req.account_str.strip()
    account_str = raw_input

    # Nếu chuỗi không chứa dấu |, thử tra cứu chuỗi đầy đủ từ SQLite DB per User Rule #4
    if '|' not in raw_input:
        db_acc = get_account_from_db(raw_input)
        if db_acc:
            account_str = db_acc
        else:
            raise HTTPException(status_code=400, detail=f"Không tìm thấy thông tin tài khoản cho key/email '{raw_input}' trong SQLite Database.")

    # Tự động lưu/cập nhật tài khoản vào SQLite DB
    save_account_to_db(account_str)

    acc_info = parse_account_string(account_str)
    username = acc_info['username']
    client_id = acc_info['client_id'] or "9e5f94bc-e8a4-4e73-b8be-63364c29d753"

    if not username:
        raise HTTPException(status_code=400, detail="Chuỗi account_str không chứa địa chỉ email hợp lệ.")

    protocol = MSGraphProtocol(api_version='v1.0')
    account = Account((client_id,), protocol=protocol, auth_flow_type='public', username=username)


    if req.use_mock:
        account.con = MagicMock()
        account.con.token_backend.has_data = True
        account.con.token_backend.token_is_expired.return_value = False
        account.con.token_backend.token_is_long_lived.return_value = True

        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            'value': [
                {
                    'id': 'msg_mock_001',
                    'subject': 'Mã xác nhận đăng nhập của bạn là 849201',
                    'bodyPreview': 'Mã OTP để hoàn tất xác minh tài khoản của bạn là 849201. Vui lòng không chia sẻ mã này.',
                    'sender': {'emailAddress': {'name': 'Security Team', 'address': 'no-reply@security.com'}},
                    'isRead': False,
                    'hasAttachments': False,
                    'createdDateTime': '2026-08-08T11:20:00Z'
                },
                {
                    'id': 'msg_mock_002',
                    'subject': 'Xác nhận đơn hàng #99482',
                    'bodyPreview': 'Đơn hàng Path Of Exile của bạn đã được xác nhận thành công.',
                    'sender': {'emailAddress': {'name': 'Xsolla Mailer', 'address': 'mailer@xsolla.com'}},
                    'isRead': True,
                    'hasAttachments': True,
                    'createdDateTime': '2026-08-07T14:30:00Z'
                }
            ]
        }
        account.con.get.return_value = mock_response
    else:
        refresh_token = acc_info.get('refresh_token')
        if refresh_token:
            import requests
            token_url = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token'
            data = {
                'client_id': client_id,
                'grant_type': 'refresh_token',
                'refresh_token': refresh_token,
                'scope': 'https://graph.microsoft.com/Mail.Read'
            }
            res = requests.post(token_url, data=data)
            if res.status_code == 200 and 'access_token' in res.json():
                access_token = res.json()['access_token']
                account.con.token_backend._cache = {'access_token': {'secret': access_token}}
                account.con.token_backend.token_is_expired = lambda username=None: False
                account.con.token_backend.token_is_long_lived = lambda username=None: True
                if account.con.session is None:
                    account.con.session = account.con.get_session(load_token=False)
                account.con.session.headers['Authorization'] = f'Bearer {access_token}'
            else:
                err_data = res.json() if res.headers.get('content-type', '').startswith('application/json') else {}
                err_msg = err_data.get('error_description', res.text)
                if 'AADSTS70000' in err_msg or 'invalid_grant' in err_data.get('error', ''):
                    # Tự động chuyển trạng thái của tài khoản sang "Hết hạn Token"
                    save_account_to_db(account_str=account_str, status="Hết hạn Token")
                    detail_msg = f"Token Outlook của tài khoản '{username}' đã bị hết hạn hoặc bị thay đổi mật khẩu từ phía Microsoft (Mã lỗi: AADSTS70000). Vui lòng cập nhật dòng token mới."
                    raise HTTPException(
                        status_code=400,
                        detail={
                            "error_code": "REFRESH_TOKEN_EXPIRED",
                            "message": detail_msg,
                            "email": username
                        }
                    )
                else:
                    detail_msg = f"Không thể lấy access_token từ refresh_token: {err_msg}"
                    raise HTTPException(
                        status_code=400,
                        detail={
                            "error_code": "O365_OAUTH_FAILED",
                            "message": detail_msg,
                            "email": username
                        }
                    )


    try:
        mailbox = account.mailbox()
        inbox = mailbox.inbox_folder()

        q = mailbox.q()
        query = None
        try:
            messages = list(inbox.get_messages(limit=req.limit, query=query, order_by='receivedDateTime desc'))
        except Exception:
            messages = list(inbox.get_messages(limit=req.limit, query=query))

        if not messages and req.keyword:
            try:
                messages = list(inbox.get_messages(limit=req.limit, order_by='receivedDateTime desc'))
            except Exception:
                messages = list(inbox.get_messages(limit=req.limit))

        if not messages:
            return OTPCodeResponse(
                status="no_messages_found",
                email=username,
                otp_code=None
            )

        parsed_email_items: List[EmailItem] = []
        for msg in messages:
            sub = getattr(msg, 'subject', '') or ''
            bp = getattr(msg, 'body_preview', str(getattr(msg, 'body', '') or '')) or ''
            snd = str(getattr(msg, 'sender', '') or '')
            dt_str = str(getattr(msg, 'created', '') or getattr(msg, 'received', '') or '')
            has_att = bool(getattr(msg, 'has_attachments', False))

            full_body = str(getattr(msg, 'body', '') or bp)
            c_sub = extract_otp_code(sub)
            c_body = extract_otp_code(full_body)
            msg_codes = list(dict.fromkeys(c_sub + c_body))

            parsed_email_items.append(EmailItem(
                id=str(getattr(msg, 'object_id', 'msg_id')),
                subject=sub,
                sender=snd,
                created_date=dt_str,
                body_preview=bp,
                body=full_body,
                has_attachments=has_att,
                otp_codes=msg_codes
            ))

        # Đảm bảo sắp xếp email theo thứ tự mới nhất xếp ở đầu
        parsed_email_items.sort(key=lambda x: x.created_date, reverse=True)

        target_msg = parsed_email_items[0]
        primary_otp = target_msg.otp_codes[0] if target_msg.otp_codes else None

        if not primary_otp:
            for item in parsed_email_items:
                if item.otp_codes:
                    primary_otp = item.otp_codes[0]
                    target_msg = item
                    break


        final_status = "success" if primary_otp else "no_otp_code_found"
        
        # Tự động cập nhật kết quả lấy mã OTP mới nhất vào duy nhất 1 bảng accounts
        save_account_to_db(
            account_str=account_str,
            otp_code=primary_otp,
            subject=target_msg.subject,
            sender=target_msg.sender
        )


        return OTPCodeResponse(
            status=final_status,
            email=username,
            otp_code=primary_otp,
            subject=target_msg.subject,
            sender=target_msg.sender,
            received_time=target_msg.created_date,
            message_preview=target_msg.body_preview,
            all_codes_found=target_msg.otp_codes,
            all_messages=parsed_email_items
        )


    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi đọc hòm thư Outlook: {str(e)}")


if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 8000))
    print(f"Đang khởi chạy API Server & Web UI Dashboard tại http://0.0.0.0:{port}...")
    uvicorn.run("api_server:app", host="0.0.0.0", port=port, reload=False)

