"""
API Server & Web UI: Dịch vụ trích xuất mã xác nhận (OTP) và hiển thị hòm thư Outlook bằng python-o365 + FastAPI.
"""

import re
import sys
import json
import sqlite3
import requests
from pathlib import Path
from typing import Optional, List, Union
from unittest.mock import MagicMock

try:
    from fastapi import FastAPI, HTTPException, Query
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
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS account_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL,
                action TEXT NOT NULL,
                details TEXT,
                user TEXT DEFAULT 'System',
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_account_logs_email ON account_logs(email)")
        conn.commit()


def add_account_log(email: str, action: str, details: str = "", user: str = "System"):
    """Ghi nhật ký hoạt động cho tài khoản email vào bảng account_logs"""
    if not email:
        return
    try:
        clean_email = email.strip().lower()
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO account_logs (email, action, details, user, timestamp)
                VALUES (?, ?, ?, ?, datetime('now'))
            """, (clean_email, action, details, user))
            conn.commit()
    except Exception as e:
        print(f"[!] Lỗi khi ghi log cho {email}: {e}")


def save_account_to_db(
    account_str: str,
    status: Optional[str] = None,
    otp_code: Optional[str] = None,
    subject: Optional[str] = None,
    sender: Optional[str] = None,
    user: Optional[str] = None,
) -> Optional[str]:
    """Lưu / Cập nhật thông tin tài khoản, danh sách trạng thái và mã OTP mới nhất vào duy nhất 1 bảng accounts trong SQLite DB"""
    if not account_str or not account_str.strip():
        return None
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

        # Kiểm tra xem email đã tồn tại chưa để ghi log đúng action
        cursor.execute("SELECT email FROM accounts WHERE LOWER(email) = LOWER(?)", (email,))
        is_existing = cursor.fetchone() is not None

        # Nếu chuỗi mới truyền vào không chứa dấu |, kiểm tra xem SQLite DB đã có chuỗi token đầy đủ chưa
        if '|' not in account_str:
            cursor.execute("SELECT account_str FROM accounts WHERE LOWER(email) = LOWER(?)", (email,))
            existing = cursor.fetchone()
            if existing and '|' in existing[0]:
                account_str = existing[0]

        cursor.execute("""
            INSERT INTO accounts (email, account_str, status, otp_code, subject, sender, updated_at)
            VALUES (?, ?, COALESCE(?, 'Chưa sử dụng'), ?, ?, ?, datetime('now'))
            ON CONFLICT(email) DO UPDATE SET
                account_str = excluded.account_str,
                status = CASE WHEN excluded.status IS NOT NULL AND excluded.status != '' THEN excluded.status ELSE accounts.status END,
                otp_code = COALESCE(excluded.otp_code, accounts.otp_code),
                subject = COALESCE(excluded.subject, accounts.subject),
                sender = COALESCE(excluded.sender, accounts.sender),
                updated_at = datetime('now')
        """, (email, account_str, status_str, otp_code, subject, sender))
        conn.commit()

    performed_by = user.strip() if user and user.strip() else "System"
    if not is_existing:
        add_account_log(
            email,
            action="IMPORT",
            details=f"Khởi tạo/Import tài khoản. Trạng thái: {status_str or 'Chưa sử dụng'}",
            user=performed_by,
        )
    elif status_str is not None:
        add_account_log(
            email,
            action="UPDATE_STATUS",
            details=f"Cập nhật trạng thái thành: {status_str}",
            user=performed_by,
        )

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
        description="Dòng thông tin tài khoản dạng: email|password|refresh_token|client_id",
    )
    keyword: Optional[str] = Field(
        default=None,
        description="Từ khóa lọc email (VD: OTP, code, mã xác nhận, Facebook...)",
    )
    limit: int = Field(default=15, ge=1, le=50, description="Số lượng email gần nhất cần quét")
    use_mock: bool = Field(
        default=False,
        description="Đặt True để chạy test dữ liệu giả lập, False để kết nối mail thật",
    )
    user: Optional[str] = Field(default=None, description="Tên người thực hiện")


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
    account_str: str = Field(
        ...,
        description="Dòng thông tin tài khoản dạng email|password|refresh_token|client_id",
    )
    status: Optional[Union[str, List[str]]] = Field(
        default=None, description="Trạng thái hoặc danh sách trạng thái của tài khoản"
    )
    user: Optional[str] = Field(default=None, description="Tên người thực hiện")


class UpdateAccountRequest(BaseModel):
    email: str = Field(..., description="Địa chỉ email tài khoản cần cập nhật")
    account_str: Optional[str] = Field(None, description="Dòng thông tin tài khoản mới")
    status: Optional[Union[str, List[str]]] = Field(
        None, description="Danh sách trạng thái hoặc chuỗi trạng thái mới"
    )
    user: Optional[str] = Field(default=None, description="Tên người thực hiện")


class BatchAddAccountsRequest(BaseModel):
    accounts: List[str] = Field(
        ..., description="Danh sách chuỗi tài khoản dạng mảng hoặc văn bản nhiều dòng"
    )
    status: Optional[Union[str, List[str]]] = Field(
        default=None, description="Trạng thái áp dụng chung cho batch"
    )
    user: Optional[str] = Field(default=None, description="Tên người thực hiện")


@app.post("/api/accounts/add")
def add_new_account(req: AddAccountRequest):
    """Thêm hoặc cập nhật tài khoản vào SQLite Database (Hỗ trợ 1 hoặc nhiều dòng)"""
    raw_str = req.account_str.strip()
    if not raw_str:
        raise HTTPException(status_code=400, detail="Vui lòng nhập chuỗi tài khoản hợp lệ.")

    lines = [l.strip() for l in raw_str.splitlines() if l.strip()]
    if len(lines) == 1:
        email = save_account_to_db(account_str=lines[0], status=req.status, user=req.user)
        if not email:
            raise HTTPException(status_code=400, detail="Chuỗi account_str không chứa địa chỉ email hợp lệ.")
        return {"status": "success", "message": f"Đã lưu tài khoản {email} vào SQLite Database", "email": email}
    else:
        return batch_add_accounts(BatchAddAccountsRequest(accounts=lines, status=req.status, user=req.user))


def process_update_account(
    target_email: Optional[str],
    account_str: Optional[str],
    status: Optional[Union[str, List[str]]],
    user: Optional[str] = None,
):
    email_clean = (target_email or "").strip()
    if not email_clean or "@" not in email_clean:
        if account_str:
            acc_info = parse_account_string(account_str)
            email_clean = acc_info.get("username", "").strip()

    if not email_clean or "@" not in email_clean:
        raise HTTPException(status_code=400, detail="Vui lòng cung cấp địa chỉ email hợp lệ.")

    existing_acc_str = get_account_from_db(email_clean)
    clean_req_str = (account_str or "").strip()

    if clean_req_str and "@" in clean_req_str:
        account_str_to_save = clean_req_str
    else:
        account_str_to_save = existing_acc_str or email_clean

    saved_email = save_account_to_db(account_str=account_str_to_save, status=status, user=user)
    if not saved_email:
        raise HTTPException(
            status_code=400, detail=f"Không thể cập nhật tài khoản {email_clean}."
        )

    return {
        "status": "success",
        "message": f"Đã cập nhật thành công tài khoản {saved_email}",
        "email": saved_email,
    }


@app.put("/api/accounts/update")
def update_account_body(req: UpdateAccountRequest):
    """Cập nhật thông tin tài khoản via Body JSON"""
    return process_update_account(
        target_email=req.email, account_str=req.account_str, status=req.status, user=req.user
    )


@app.put("/api/accounts/{email:path}")
def update_account_path(email: str, req: UpdateAccountRequest):
    """Cập nhật thông tin tài khoản via Path Email"""
    clean_email = email if email and email != "update" else None
    return process_update_account(
        target_email=clean_email or req.email,
        account_str=req.account_str,
        status=req.status,
        user=req.user,
    )


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
            email = save_account_to_db(account_str=line, status=req.status, user=req.user)
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


class ResetAllStatusRequest(BaseModel):
    status: str = Field(default="Chưa sử dụng", description="Trạng thái áp dụng cho toàn bộ tài khoản")
    user: Optional[str] = Field(default=None, description="Tên người thực hiện")


@app.post("/api/accounts/reset-all-status")
def reset_all_accounts_status(req: ResetAllStatusRequest):
    """Cập nhật trạng thái cho TOÀN BỘ tài khoản trong SQLite DB về trạng thái chỉ định (mặc định: Chưa sử dụng)"""
    new_status = req.status.strip() if req.status and req.status.strip() else "Chưa sử dụng"
    performed_by = req.user.strip() if req.user and req.user.strip() else "System"

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT email FROM accounts")
        all_emails = [r[0] for r in cursor.fetchall()]

        cursor.execute("UPDATE accounts SET status = ?, updated_at = datetime('now')", (new_status,))
        updated_count = cursor.rowcount
        conn.commit()

    for email in all_emails:
        add_account_log(
            email,
            action="UPDATE_STATUS",
            details=f"Chuyển trạng thái hàng loạt thành: {new_status}",
            user=performed_by,
        )

    return {
        "status": "success",
        "message": f"Đã chuyển trạng thái của toàn bộ {updated_count} tài khoản về '{new_status}'",
        "updated_count": updated_count,
    }


class BoosterLinkItem(BaseModel):
    email: str = Field(..., description="Địa chỉ email hoặc account string")
    game_title: Optional[str] = Field(None, description="Tên game (VD: Diablo IV, PoE...)")


class BatchSyncBoosterLinksRequest(BaseModel):
    items: List[BoosterLinkItem] = Field(..., description="Danh sách tài khoản booster")
    user: Optional[str] = Field(default=None, description="Tên người thực hiện")


@app.post("/api/accounts/sync-booster-links")
def sync_booster_links(req: BatchSyncBoosterLinksRequest):
    """Đồng bộ tự động trạng thái email từ danh sách Booster Accounts (Diablo -> đã mua game diablo, PoE 1 -> đã mua rương poe1 cho trader, PoE 2 -> đã mua rương poe2 cho trader)
    LƯU Ý: Không ghi đè nếu tài khoản đã ở trạng thái 'Hết hạn Token' hoặc 'Khóa/Ban'.
    """
    updated_count = 0
    performed_by = req.user.strip() if req.user and req.user.strip() else "System"

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()

        for item in req.items:
            raw_email = (item.email or "").strip()
            if not raw_email:
                continue

            found_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', raw_email)
            if not found_emails:
                continue

            game_title = (item.game_title or "").lower()

            target_status = None
            if "poe 2" in game_title or "poe2" in game_title or "path of exile 2" in game_title:
                target_status = "đã mua rương poe2 cho trader"
            elif "poe" in game_title or "poe1" in game_title or "path of exile" in game_title:
                target_status = "đã mua rương poe1 cho trader"
            elif "diablo" in game_title or "d4" in game_title:
                target_status = "đã mua game diablo"

            if target_status:
                for email_clean in found_emails:
                    email_clean = email_clean.lower().strip()

                    # Kiểm tra trạng thái hiện tại trong DB:
                    cursor.execute("SELECT status FROM accounts WHERE LOWER(email) = LOWER(?)", (email_clean,))
                    existing = cursor.fetchone()
                    if existing and existing[0]:
                        curr_st = existing[0].lower()
                        # 1. Không ghi đè nếu tài khoản đã bị Ban / Khóa
                        if "ban" in curr_st or "khóa" in curr_st:
                            continue
                        # 2. Ưu tiên PoE 2: Nếu DB đang là PoE 2 thì không hạ xuống PoE 1 hay Diablo
                        if "poe2" in curr_st or "poe 2" in curr_st or "rương poe2" in curr_st:
                            if target_status != "đã mua rương poe2 cho trader":
                                continue
                        # 3. Ưu tiên PoE 1: Nếu DB đang là PoE 1 thì không hạ xuống Diablo
                        if ("poe" in curr_st or "poe1" in curr_st) and "poe2" not in curr_st:
                            if target_status == "đã mua game diablo":
                                continue

                    saved = save_account_to_db(account_str=email_clean, status=target_status, user=performed_by)
                    if saved:
                        updated_count += 1

    return {
        "status": "success",
        "message": f"Đã đồng bộ thành công trạng thái cho {updated_count} email dựa trên Game Title.",
        "updated_count": updated_count,
    }


class EmailOnlyRequest(BaseModel):
    email: str = Field(..., description="Địa chỉ email tài khoản Outlook (VD: sylvesterrojas997795@outlook.com)")
    keyword: Optional[str] = Field(None, description="Từ khóa lọc email")
    limit: Optional[int] = Field(15, description="Số lượng email tối đa cần đọc")
    user: Optional[str] = Field(None, description="Tên người thực hiện")


@app.get("/api/accounts/{email:path}/logs")
def get_account_logs(email: str):
    """Lấy danh sách nhật ký hoạt động của tài khoản email (Per Rule #5 Remote Log API)"""
    clean_email = email.strip().lower()
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT id, email, action, details, user, timestamp 
            FROM account_logs 
            WHERE LOWER(email) = LOWER(?) 
            ORDER BY id DESC LIMIT 100
        """,
            (clean_email,),
        )
        rows = cursor.fetchall()
        logs = [
            {
                "id": r[0],
                "email": r[1],
                "action": r[2],
                "details": r[3],
                "user": r[4],
                "timestamp": r[5],
            }
            for r in rows
        ]
        return {"status": "success", "email": clean_email, "logs": logs}


def auto_renew_outlook_token(account_str: str, user: Optional[str] = None) -> Optional[str]:
    """
    Tự động gọi API Renew Outlook Token (https://api-tools.shopmailmmo.com/api/v1/public/outlook/renew)
    khi refresh token bị hết hạn. Trả về account_str mới nếu thành công, hoặc None nếu thất bại.
    """
    if not account_str or '|' not in account_str:
        return None

    renew_url = "https://api-tools.shopmailmmo.com/api/v1/public/outlook/renew"
    try:
        res = requests.post(renew_url, json={"data": account_str}, timeout=15)
        if res.status_code == 200:
            res_json = res.json()
            if res_json.get("success") and res_json.get("data"):
                new_account_str = res_json["data"]
                # Lưu account_str mới và chuyển trạng thái về "Chưa sử dụng"
                save_account_to_db(account_str=new_account_str, status="Chưa sử dụng", user=user)
                acc_info = parse_account_string(new_account_str)
                email = acc_info.get("username")
                if email:
                    add_account_log(
                        email,
                        action="AUTO_RENEW_TOKEN",
                        details="Tự động Gia Hạn (Renew) Token thành công qua api-tools.shopmailmmo.com",
                        user=user or "System",
                    )
                return new_account_str
            else:
                print(f"[AUTO_RENEW] API shopmailmmo báo lỗi: {res_json}")
    except Exception as e:
        print(f"[AUTO_RENEW_EXCEPTION] Lỗi khi kết nối tới api-tools.shopmailmmo.com: {e}")
    return None


class RenewTokenRequest(BaseModel):
    email: str = Field(..., description="Địa chỉ email hoặc account string")
    user: Optional[str] = Field(default=None, description="Tên người thực hiện")


@app.post("/api/accounts/renew-token")
def manual_renew_token_endpoint(req: RenewTokenRequest):
    """API gia hạn (renew) Token thủ công cho 1 tài khoản email qua shopmailmmo"""
    raw_input = req.email.strip()
    account_str = raw_input
    if '|' not in raw_input:
        db_acc = get_account_from_db(raw_input)
        if db_acc:
            account_str = db_acc
        else:
            raise HTTPException(
                status_code=400, detail=f"Không tìm thấy token của email '{raw_input}' trong SQLite DB"
            )

    renewed_str = auto_renew_outlook_token(account_str, user=req.user)
    if renewed_str:
        return {
            "status": "success",
            "message": f"Đã gia hạn (Renew) Token thành công cho tài khoản '{req.email}'!",
            "account_str": renewed_str,
        }
    else:
        raise HTTPException(
            status_code=400,
            detail=f"Không thể gia hạn (Renew) Token cho '{req.email}' qua api-tools.shopmailmmo.com. Vui lòng kiểm tra lại dòng token.",
        )


@app.delete("/api/accounts/delete/{email:path}")
@app.delete("/api/accounts/{email:path}")
def delete_account(email: str, user: Optional[str] = None):
    """Xóa tài khoản khỏi SQLite Database theo email"""
    if email in ("batch", "add"):
        raise HTTPException(status_code=400, detail="Tên email không hợp lệ.")
    success = delete_account_from_db(email)
    if not success:
        raise HTTPException(
            status_code=404,
            detail=f"Không tìm thấy tài khoản '{email}' trong SQLite Database.",
        )
    performed_by = user.strip() if user and user.strip() else "System"
    add_account_log(email, action="DELETE", details="Xóa tài khoản khỏi SQLite Database", user=performed_by)
    return {"status": "success", "message": f"Đã xóa tài khoản {email} khỏi SQLite Database"}


@app.get("/api/get-code-by-email")
def get_code_by_email_get(email: str = Query(..., description="Địa chỉ email tài khoản Outlook")):
    """
    Endpoint tra cứu mã OTP nhanh dạng GET.
    Ví dụ: GET /api/get-code-by-email?email=sylvesterrojas997795@outlook.com
    """
    code_req = AccountCodeRequest(account_str=email)
    return get_verification_code(code_req)


@app.post("/api/get-code-by-email")
def get_code_by_email_post(req: EmailOnlyRequest):
    """
    Endpoint tra cứu mã OTP nhanh dạng POST.
    Ví dụ: POST /api/get-code-by-email với body {"email": "sylvesterrojas997795@outlook.com"}
    """
    code_req = AccountCodeRequest(account_str=req.email, keyword=req.keyword)
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
        if db_acc and '|' in db_acc:
            account_str = db_acc
        else:
            raise HTTPException(
                status_code=400,
                detail=f"Tài khoản '{raw_input}' chưa có Script Key (chuỗi token đầy đủ email|password|refresh_token|client_id). Vui lòng bổ sung Script Key.",
            )

    # Tự động lưu/cập nhật tài khoản vào SQLite DB
    save_account_to_db(account_str, user=req.user)

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
            access_token = None

            if res.status_code == 200 and 'access_token' in res.json():
                access_token = res.json()['access_token']
            else:
                # Tự động gia hạn (Renew) token 1 lần qua API shopmailmmo khi bị hết hạn
                renewed_str = auto_renew_outlook_token(account_str, user=req.user)
                if renewed_str:
                    renewed_info = parse_account_string(renewed_str)
                    new_refresh = renewed_info.get('refresh_token')
                    new_client_id = renewed_info.get('client_id') or client_id
                    if new_refresh:
                        data['refresh_token'] = new_refresh
                        data['client_id'] = new_client_id
                        res_retry = requests.post(token_url, data=data)
                        if res_retry.status_code == 200 and 'access_token' in res_retry.json():
                            access_token = res_retry.json()['access_token']
                            account_str = renewed_str
                            acc_info = renewed_info

            if access_token:
                account.con.token_backend._cache = {'access_token': {'secret': access_token}}
                account.con.token_backend.token_is_expired = lambda username=None: False
                account.con.token_backend.token_is_long_lived = lambda username=None: True
                if account.con.session is None:
                    account.con.session = account.con.get_session(load_token=False)
                account.con.session.headers['Authorization'] = f'Bearer {access_token}'
            else:
                save_account_to_db(account_str=account_str, status="Hết hạn Token", user=req.user)
                detail_msg = f"Token Outlook của tài khoản '{username}' đã bị hết hạn và không thể tự động gia hạn (Renew). Vui lòng cập nhật dòng token mới."
                raise HTTPException(
                    status_code=400,
                    detail={
                        "error_code": "REFRESH_TOKEN_EXPIRED",
                        "message": detail_msg,
                        "email": username,
                    },
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
        
        performed_by = req.user.strip() if req.user and req.user.strip() else "System"
        # Tự động cập nhật kết quả lấy mã OTP mới nhất vào duy nhất 1 bảng accounts
        save_account_to_db(
            account_str=account_str,
            otp_code=primary_otp,
            subject=target_msg.subject,
            sender=target_msg.sender,
            user=performed_by,
        )
        add_account_log(
            username,
            action="FETCH_OTP",
            details=f"Đã đọc hòm thư & lấy mã OTP: {primary_otp or 'Không tìm thấy mã'} | Tiêu đề: {target_msg.subject or 'N/A'}",
            user=performed_by,
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


class SettingsPayload(BaseModel):
    allowed_view_roles: Optional[List[str]] = None
    allowed_manage_roles: Optional[List[str]] = None
    custom_statuses: Optional[List[dict]] = None


@app.get("/api/settings")
def get_settings():
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT key, value FROM settings")
        rows = cursor.fetchall()
        result = {}
        for k, v in rows:
            try:
                result[k] = json.loads(v)
            except Exception:
                result[k] = v

        default_roles = [
            "System Manager",
            "Administrator",
        ]
        default_statuses = [
            {"name": "Hoạt động", "icon": "✅", "color": "emerald"},
            {"name": "Đang dùng", "icon": "🔥", "color": "indigo"},
            {"name": "Dự phòng", "icon": "📦", "color": "sky"},
            {"name": "Tài khoản chính", "icon": "⭐", "color": "purple"},
            {"name": "Hết hạn Token", "icon": "⚠️", "color": "amber"},
            {"name": "Khóa/Ban", "icon": "🚫", "color": "red"},
        ]
        return {
            "status": "success",
            "allowed_view_roles": result.get("allowed_view_roles", default_roles),
            "allowed_manage_roles": result.get("allowed_manage_roles", default_roles),
            "custom_statuses": result.get("custom_statuses", default_statuses),
        }


@app.post("/api/settings")
def save_settings(payload: SettingsPayload):
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        if payload.allowed_view_roles is not None:
            cursor.execute(
                "INSERT INTO settings (key, value, updated_at) VALUES ('allowed_view_roles', ?, datetime('now')) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=datetime('now')",
                (json.dumps(payload.allowed_view_roles),),
            )
        if payload.allowed_manage_roles is not None:
            cursor.execute(
                "INSERT INTO settings (key, value, updated_at) VALUES ('allowed_manage_roles', ?, datetime('now')) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=datetime('now')",
                (json.dumps(payload.allowed_manage_roles),),
            )
        if payload.custom_statuses is not None:
            cursor.execute(
                "INSERT INTO settings (key, value, updated_at) VALUES ('custom_statuses', ?, datetime('now')) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=datetime('now')",
                (json.dumps(payload.custom_statuses),),
            )
            # Tự động dọn dẹp các trạng thái bị xóa khỏi tất cả tài khoản trong bảng accounts DB
            valid_names = set(
                s.get("name", "").strip() for s in payload.custom_statuses if s.get("name")
            )
            cursor.execute(
                "SELECT email, status FROM accounts WHERE status IS NOT NULL AND status != ''"
            )
            all_accs = cursor.fetchall()
            for acc_email, acc_status in all_accs:
                tags = [t.strip() for t in acc_status.split(",") if t.strip()]
                cleaned_tags = [t for t in tags if t in valid_names]
                new_status_str = ", ".join(cleaned_tags) if cleaned_tags else "Hoạt động"
                if new_status_str != acc_status:
                    cursor.execute(
                        "UPDATE accounts SET status = ?, updated_at = datetime('now') WHERE email = ?",
                        (new_status_str, acc_email),
                    )
        conn.commit()
    return {
        "status": "success",
        "message": "Đã lưu cấu hình phân quyền và trạng thái tập trung vào SQLite DB",
    }


if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 8005))
    print(f"Đang khởi chạy API Server & Web UI Dashboard tại http://0.0.0.0:{port}...")
    uvicorn.run("api_server:app", host="0.0.0.0", port=port, reload=False)

