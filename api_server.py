"""
API Server & Web UI: Dịch vụ trích xuất mã xác nhận (OTP) và hiển thị hòm thư Outlook bằng python-o365 + FastAPI.
"""

import re
import sys
from pathlib import Path
from typing import Optional, List
from unittest.mock import MagicMock

try:
    from fastapi import FastAPI, HTTPException
    from fastapi.staticfiles import StaticFiles
    from fastapi.responses import FileResponse
    from pydantic import BaseModel, Field
    import uvicorn
    from O365 import Account, MSGraphProtocol
except ModuleNotFoundError as err:
    print(f"[!] Thiếu thư viện phụ thuộc ({err}). Vui lòng chạy lệnh:")
    print("    pip install fastapi uvicorn O365 msal requests_oauthlib beautifulsoup4 python-dateutil")
    sys.exit(1)

app = FastAPI(
    title="Outlook Mail & OTP Extractor Dashboard",
    description="API Server và Giao diện Web lấy mã OTP 1-click & xem toàn bộ hòm thư Outlook.",
    version="2.0.0"
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
    if rf.endswith('$$'):
        rf = rf[:-1]
    return {
        'username': parts[0] if len(parts) > 0 else '',
        'password': parts[1] if len(parts) > 1 else '',
        'refresh_token': rf,
        'client_id': parts[3] if len(parts) > 3 else ''
    }


def extract_otp_code(text: str) -> List[str]:
    """Trích xuất danh sách mã xác nhận (4-8 chữ số) từ văn bản"""
    if not text:
        return []
    cleaned_text = re.sub(r'\b20[12]\d\b', '', text)
    matches = re.findall(r'\b\d{4,8}\b', cleaned_text)
    seen = set()
    unique_codes = []
    for code in matches:
        if code not in seen:
            seen.add(code)
            unique_codes.append(code)
    return unique_codes


@app.get("/")
def read_root():
    """Phục vụ trang Giao diện Web UI Dashboard"""
    index_file = static_dir / "index.html"
    if index_file.exists():
        return FileResponse(index_file)
    return {"message": "Web UI chưa khởi tạo. Vui lòng tạo static/index.html"}


@app.get("/health")
def health_check():
    """Endpoint kiểm tra sức khỏe của API Server"""
    return {"status": "ok", "service": "Outlook OTP Extractor API Server"}


@app.post("/api/get-code", response_model=OTPCodeResponse)
def get_verification_code(req: AccountCodeRequest):
    """
    Nhận chuỗi tài khoản email|password|refresh_token|client_id,
    đọc hòm thư Outlook và trả về mã OTP cùng toàn bộ danh sách email.
    """
    acc_info = parse_account_string(req.account_str)
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
                err_msg = res.json().get('error_description', res.text)
                raise HTTPException(status_code=400, detail=f"Không thể lấy access_token từ refresh_token: {err_msg}")

    try:
        mailbox = account.mailbox()
        inbox = mailbox.inbox_folder()

        q = mailbox.q()
        query = None
        if req.keyword:
            query = q.contains('subject', req.keyword) | q.contains('body', req.keyword)

        messages = list(inbox.get_messages(limit=req.limit, query=query))

        if not messages and req.keyword:
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
            bp = getattr(msg, 'body_preview', str(msg.body or '')) or ''
            snd = str(getattr(msg, 'sender', '') or '')
            dt_str = str(getattr(msg, 'created', '') or '')
            has_att = bool(getattr(msg, 'has_attachments', False))

            c_sub = extract_otp_code(sub)
            c_body = extract_otp_code(bp)
            msg_codes = list(dict.fromkeys(c_sub + c_body))

            full_body = str(getattr(msg, 'body', '') or bp)
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


        target_msg = parsed_email_items[0]
        primary_otp = target_msg.otp_codes[0] if target_msg.otp_codes else None

        if not primary_otp:
            for item in parsed_email_items:
                if item.otp_codes:
                    primary_otp = item.otp_codes[0]
                    target_msg = item
                    break

        return OTPCodeResponse(
            status="success" if primary_otp else "no_otp_code_found",
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
    print("Đang khởi chạy API Server & Web UI Dashboard tại http://0.0.0.0:8000...")
    uvicorn.run("api_server:app", host="0.0.0.0", port=8000, reload=False)
