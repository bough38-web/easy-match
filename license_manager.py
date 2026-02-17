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

# --- Developer Exemption ---
DEVELOPER_HWIDS = ["9B3B04F2D13FFED1"]

def _check_persistent_trial():
    """
    Checks for a hidden persistent file in user home to prevent trial abuse.
    Returns: (is_allowed: bool, message: str)
    """
    try:
        # Developer Bypass
        from security_utils import get_hwid
        current_hwid = get_hwid()
        if current_hwid in DEVELOPER_HWIDS:
            return True, "Developer Device (Bypassed)"

        home = os.path.expanduser("~")
        trial_file = os.path.join(home, ".excelmatcher_trial")
        today = datetime.date.today()

        if not os.path.exists(trial_file):
            # First run
            try:
                with open(trial_file, "w") as f:
                    f.write(today.strftime("%Y-%m-%d"))
                # Hide file on Windows
                if os.name == 'nt':
                    import ctypes
                    try:
                        ctypes.windll.kernel32.SetFileAttributesW(trial_file, 2) # Hidden
                    except:
                        pass
            except:
                pass # Permission error?
            return True, "First Run"
        
        # Check existing
        with open(trial_file, "r") as f:
            first_run_s = f.read().strip()
        
        try:
            first_run = datetime.datetime.strptime(first_run_s, "%Y-%m-%d").date()
        except ValueError:
            # Date corrupted? Reset to today to be safe (or block?)
            # Blocking might affect innocent users if file gets corrupted.
            # Let's assume valid if corrupted but recommend reinstall.
            return True, "Date Error"

        days_diff = (today - first_run).days
        if days_diff > 30:
            return False, f"무료 체험 기간(30일)이 만료되었습니다. (최초 실행: {first_run_s})"
        
        return True, f"체험 기간 중 ({days_diff}/30일)"

    except Exception as e:
        # Fail safe
        return True, str(e)


def ensure_license():
    """
    1-month trial for new users, then registration required.
    Includes persistent anti-abuse check.
    """
    import tkinter as tk
    from tkinter import messagebox, simpledialog
    from security_utils import get_hwid

    ok, msg, info = validate_license()
    
    # [Anti-Abuse] Persistent Check
    # Even if license is valid, if it's a TRIAL license, check persistence.
    # If it's a PAID license (personal/business), skip persistence check.
    
    is_trial_key = (ok and info and info.get("type", "personal") == "trial")
    no_key = (not ok)

    if is_trial_key or no_key:
        p_ok, p_msg = _check_persistent_trial()
        if not p_ok:
            # Determine if we should block
            # If it's a valid trial key but persistent says expired -> Block
            # If no key and persistent says expired -> Block
            ok = False
            msg = p_msg # Override error message
    
    if ok:
        return True, info
        
    # (1) 만약 라이선스 파일이 없고 + Persistent도 OK라면 -> 1개월 체험판 자동 생성
    if not os.path.exists(LICENSE_FILE):
        # Double check persistence
        p_ok, _ = _check_persistent_trial()
        if p_ok:
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
                # Topmost to ensure visibility
                root.attributes("-topmost", True)
                messagebox.showinfo("체험판 시작", f"이지 매치 v1.0.0에 오신 것을 환영합니다!\n\n1개월 무료 체험판이 활성화되었습니다.\n만료일: {expiry_str}")
                root.destroy()
                return True, info

    # (2) 라이선스가 만료되었거나 기기가 바뀌었을 경우 -> 등록 UI
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    
    error_msg = f"라이선스 확인 실패: {msg}\n\n정품 라이선스 키를 입력하시겠습니까?\n(구매 문의: bough38@gmail.com)"
    choice = messagebox.askyesno("라이선스 인증 필요", error_msg)
    
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

