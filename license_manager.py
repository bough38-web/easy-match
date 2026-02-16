import os
import datetime
import json
from config import LICENSE_FILE
try:
    from license_key import generate_key, validate_key
except ImportError:
    # Fallback if license_key.py is missing (e.g. during dev)
    def generate_key(*args): return ""
    def validate_key(*args): return False, {}

def load_license_key():
    if not os.path.exists(LICENSE_FILE):
        return None
    try:
        with open(LICENSE_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            # If it looks like JSON, it's legacy
            if content.startswith("{") and content.endswith("}"):
                return _migrate_legacy_license(content)
            return content
    except Exception:
        return None

def _migrate_legacy_license(json_content):
    """
    Attempt to migrate legacy JSON license to new Key format.
    If valid, generate key and save. If invalid/expired, return None (force new auth).
    """
    try:
        data = json.loads(json_content)
        expiry_s = data.get("expiry")
        l_type = data.get("type", "personal")
        
        # Check if expired
        try:
            expiry = datetime.datetime.strptime(expiry_s, "%Y-%m-%d").date()
            if expiry < datetime.date.today():
                return None # Expired, force new key
        except:
            return None # Invalid date
            
        # Generate new key
        new_key = generate_key(expiry_s, l_type)
        
        # Save new key overwriting the file
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            f.write(new_key)
            
        return new_key
    except:
        return None

def load_license():
    """
    Backward compatibility wrapper for AdminPanel.
    Returns dict or None.
    """
    key_str = load_license_key()
    if not key_str:
        return None
    valid, info = validate_key(key_str)
    if valid:
        return info
    return None

def validate_license():
    key_str = load_license_key()
    if not key_str:
        return False, "라이선스 키 파일 없음", None

    is_valid, data = validate_key(key_str)
    if not is_valid:
        return False, "유효하지 않은 라이선스 키", None

    # Date Check
    try:
        expiry_s = data.get("expiry")
        expiry = datetime.datetime.strptime(expiry_s, "%Y-%m-%d").date()
        today = datetime.date.today()
        
        if today > expiry:
            return False, f"라이선스 만료 ({expiry_s})", None
            
        # HWID Check
        hwid_in_key = data.get("hwid")
        if hwid_in_key:
            from security_utils import get_hwid
            current_hwid = get_hwid()
            if hwid_in_key != current_hwid:
                return False, f"등록되지 않은 기기입니다. (ID: {current_hwid})", data

        # Remote Kill-Switch Check
        from security_utils import check_remote_block
        is_blocked, block_msg = check_remote_block(license_key=load_license_key())
        if is_blocked:
            return False, block_msg, data

        # Send usage telemetry (Silent)
        from security_utils import send_usage_log
        send_usage_log(data, action="Validate")
            
        return True, "OK", data
    except Exception as e:
        return False, f"데이터 파싱 오류: {e}", None

def save_license(expiry: str, lic_type: str, hwid: str = None):
    """
    Generates a new key and saves it to file.
    """
    new_key = generate_key(expiry, lic_type, hwid=hwid)
    save_license_key(new_key)
    return {"expiry": expiry, "type": lic_type, "key": new_key, "hwid": hwid}

def save_license_key(key_str: str):
    """
    Saves the raw key string to the license file.
    """
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        f.write(key_str.strip())

def ensure_license():
    """
    1-month trial for new users, then registration required.
    """
    import tkinter as tk
    from tkinter import messagebox, simpledialog
    from security_utils import get_hwid

    ok, msg, info = validate_license()
    if ok:
        return True, info
        
    # (1) 만약 라이선스 파일이 아예 없다면 -> 1개월 체험판 자동 생성
    if not os.path.exists(LICENSE_FILE):
        expiry_date = datetime.date.today() + datetime.timedelta(days=30)
        expiry_str = expiry_date.strftime("%Y-%m-%d")
        hwid = get_hwid()
        
        save_license(expiry_str, "trial", hwid=hwid)
        
        # 다시 검증
        ok, msg, info = validate_license()
        if ok:
            # 첫 실행 시에만 환영 메시지 (선택 사항)
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("체험판 시작", f"이지 매치 v1.0.0에 오신 것을 환영합니다!\n\n1개월 무료 체험판이 활성화되었습니다.\n만료일: {expiry_str}")
            root.destroy()
            return True, info

    # (2) 라이선스가 만료되었거나 기기가 바뀌었을 경우 -> 등록 UI
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    
    error_msg = f"라이선스 확인 실패: {msg}\n\n제품 키를 입력하시겠습니까?"
    choice = messagebox.askyesno("라이선선 인증 필요", error_msg)
    
    if choice:
        # HWID 안내 포함
        curr_hwid = get_hwid()
        prompt = f"현재 기기 ID: {curr_hwid}\n\n구매하신 제품 키를 입력하세요:"
        user_key = simpledialog.askstring("제품 키 등록", prompt)
        
        if user_key and user_key.strip():
            save_license_key(user_key)
            v_ok, v_msg, v_info = validate_license()
            if v_ok:
                messagebox.showinfo("성공", "정식 라이선스가 등록되었습니다.")
                root.destroy()
                return True, v_info
            else:
                messagebox.showerror("실패", f"유효하지 않은 키입니다.\n{v_msg}")
        
    root.destroy()
    return False, "Unauthorized"

