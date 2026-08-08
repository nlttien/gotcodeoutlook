"""
API Server: Dịch vụ trích xuất mã xác nhận (OTP / Verification Code) từ hòm thư Outlook bằng python-o365.
"""

import re
import sys
from typing import Optional, List
from unittest.mock import MagicMock

try:
    from fastapi import FastAPI, HTTPException
    from pydantic import BaseModel, Field
    import uvicorn
    from O365 import Account, MSGraphProtocol
except ModuleNotFoundError as err:
    print(f"[!] Thiếu thư viện phụ thuộc ({err}). Vui lòng chạy lệnh:")
    print("    pip install fastapi uvicorn O365 msal requests_oauthlib beautifulsoup4 python-dateutil")
    sys.exit(1)

app = FastAPI(
    title="Outlook OTP Extractor API",
    description="API nhận dòng thông tin tài khoản Outlook và trích xuất mã xác nhận (OTP / Code) mới nhất.",
    version="1.0.0"
)


class AccountCodeRequest(BaseModel):
    account_str: str = Field(
        ...,
        description="Dòng thông tin tài khoản dạng: email|password|refresh_token|client_id",
        example="christinakazunas1125@outlook.com|ChristinaKazunas694|M.C542_BAY...|9e5f94bc-e8a4-4e73-b8be-63364c29d753"
    )
    keyword: Optional[str] = Field(
        default=None,
        description="Từ khóa lọc email (VD: OTP, code, mã xác nhận, Facebook, Google...)"
    )
    limit: int = Field(default=5, ge=1, le=20, description="Số lượng email gần nhất cần quét")
    use_mock: bool = Field(default=False, description="Đặt True để chạy test dữ liệu giả lập, False để kết nối mail thật")


class OTPCodeResponse(BaseModel):
    status: str
    email: str
    otp_code: Optional[str] = None
    subject: Optional[str] = None
    sender: Optional[str] = None
    received_time: Optional[str] = None
    message_preview: Optional[str] = None
    all_codes_found: List[str] = []


def parse_account_string(account_str: str) -> dict:
    """Bóc tách dòng thông tin tài khoản định dạng: email|password|refresh_token|client_id"""
    parts = [p.strip() for p in account_str.split('|')]
    return {
        'username': parts[0] if len(parts) > 0 else '',
        'password': parts[1] if len(parts) > 1 else '',
        'refresh_token': parts[2] if len(parts) > 2 else '',
        'client_id': parts[3] if len(parts) > 3 else ''
    }


def extract_otp_code(text: str) -> List[str]:
    """Trích xuất danh sách mã xác nhận (4-8 chữ số) từ văn bản"""
    if not text:
        return []
    
    # Loại bỏ năm 202X, 201X để tránh nhầm năm với mã OTP
    cleaned_text = re.sub(r'\b20[12]\d\b', '', text)
    
    # Tìm tất cả các dãy số từ 4 đến 8 chữ số
    matches = re.findall(r'\b\d{4,8}\b', cleaned_text)
    
    # Loại bỏ trùng lặp và giữ nguyên thứ tự xuất hiện
    seen = set()
    unique_codes = []
    for code in matches:
        if code not in seen:
            seen.add(code)
            unique_codes.append(code)
            
    return unique_codes


@app.get("/health")
def health_check():
    """Endpoint kiểm tra sức khỏe của API Server"""
    return {"status": "ok", "service": "Outlook OTP Extractor API"}


@app.post("/api/get-code", response_model=OTPCodeResponse)
def get_verification_code(req: AccountCodeRequest):
    """
    Nhận chuỗi tài khoản email|password|refresh_token|client_id,
    đọc hòm thư Outlook và trả về mã xác nhận (OTP) mới nhất.
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
                }
            ]
        }
        account.con.get.return_value = mock_response

    try:
        mailbox = account.mailbox()
        inbox = mailbox.inbox_folder()

        q = mailbox.q()
        query = None
        if req.keyword:
            query = q.contains('subject', req.keyword) | q.contains('body', req.keyword)

        messages = list(inbox.get_messages(limit=req.limit, query=query))

        if not messages:
            return OTPCodeResponse(
                status="no_messages_found",
                email=username,
                otp_code=None
            )

        target_msg = messages[0]
        subject = getattr(target_msg, 'subject', '')
        body_preview = getattr(target_msg, 'body_preview', str(target_msg.body))
        sender = str(getattr(target_msg, 'sender', ''))
        received_time = str(getattr(target_msg, 'created', ''))

        codes_in_subject = extract_otp_code(subject)
        codes_in_body = extract_otp_code(body_preview)

        all_codes = list(dict.fromkeys(codes_in_subject + codes_in_body))
        primary_otp = all_codes[0] if all_codes else None

        return OTPCodeResponse(
            status="success" if primary_otp else "no_otp_code_found",
            email=username,
            otp_code=primary_otp,
            subject=subject,
            sender=sender,
            received_time=received_time,
            message_preview=body_preview,
            all_codes_found=all_codes
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi đọc hòm thư Outlook: {str(e)}")


if __name__ == '__main__':
    print("Đang khởi chạy API Server tại http://0.0.0.0:8000...")
    uvicorn.run("api_server:app", host="0.0.0.0", port=8000, reload=False)
