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
            
        return True, "OK", data
    except Exception as e:
        return False, f"데이터 파싱 오류: {e}", None

def save_license(expiry: str, lic_type: str):
    """
    Generates a new key and saves it to file.
    """
    new_key = generate_key(expiry, lic_type)
    save_license_key(new_key)
    return {"expiry": expiry, "type": lic_type, "key": new_key}

def save_license_key(key_str: str):
    """
    Saves the raw key string to the license file.
    """
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        f.write(key_str.strip())

def ensure_license():
    """
    Create a default personal license if none exists.
    """
    import tkinter as tk
    from tkinter import messagebox, simpledialog

    ok, msg, info = validate_license()
    if ok:
        return True, info
        
    # UI for Key Input or Creation
    root = tk.Tk()
    root.withdraw()

    # If file exists but invalid -> Ask for key
    # If file not exists -> Ask create or key
    
    msg = "유효한 라이선스가 없습니다.\n\n[예] 무료 개인 라이선스 생성 (1년)\n[아니오] 제품 키 직접 입력"
    choice = messagebox.askyesno("라이선스 등록", msg)
    
    if choice:
        # Create Personal
        next_year = datetime.date.today().replace(year=datetime.date.today().year + 1)
        expiry_str = next_year.strftime("%Y-%m-%d")
        save_license(expiry_str, "personal")
        
        messagebox.showinfo("완료", f"개인 라이선스가 생성되었습니다.\n만료: {expiry_str}")
        root.destroy()
        return True, {"type": "personal", "expiry": expiry_str}
    else:
        # Input Key
        user_key = simpledialog.askstring("제품 키 입력", "보유하신 제품 키를 입력하세요:\n(예: EM-XXXX...)")
        if user_key and user_key.strip():
            # Save raw key
            with open(LICENSE_FILE, "w", encoding="utf-8") as f:
                f.write(user_key.strip())
            
            # Re-validate
            v_ok, v_msg, v_info = validate_license()
            if v_ok:
                messagebox.showinfo("성공", "라이선스가 정상 등록되었습니다.")
                root.destroy()
                return True, v_info
            else:
                messagebox.showerror("실패", f"라이선스 키가 유효하지 않습니다.\n{v_msg}")
                root.destroy()
                return False, {}
        else:
            root.destroy()
            return False, "Cancelled"

