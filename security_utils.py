import os
import sys
import uuid
import platform
import subprocess
import json
import threading
import requests

def get_hwid():
    """
    Returns a unique hardware ID for the machine.
    Combines machine node and OS specific identifiers.
    """
    try:
        # 1. Basic node ID (MAC address based)
        node = hex(uuid.getnode())
        
        # 2. OS Specific Serials
        os_serial = ""
        if sys.platform == "darwin":
            # macOS: IORegistry logic for IOPlatformSerialNumber
            cmd = "ioreg -l | grep IOPlatformSerialNumber | awk '{print $4}' | sed 's/\"//g'"
            os_serial = subprocess.check_output(cmd, shell=True).decode().strip()
        elif sys.platform == "win32":
            # Windows: wmic logic for UUID
            cmd = "wmic csproduct get uuid"
            os_serial = subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip()
        
        # Combine and hash for a cleaner string
        raw = f"{node}-{os_serial}-{platform.processor()}"
        import hashlib
        return hashlib.sha1(raw.encode()).hexdigest()[:16].upper()
    except:
        # Fallback to just node if subprocess fails
        return hex(uuid.getnode())[2:].upper()

def send_usage_log(license_info, action="Launch"):
    """
    Silently sends a usage log to a pre-defined webhook.
    Logic is wrapped in a thread to avoid UI lag.
    """
    # Note: The user can provide a real webhook URL later. 
    # For now, we use a placeholder or a default tracking endpoint if provided.
    WEBHOOK_URL = os.environ.get("EM_TRACKING_WEBHOOK", "")
    if not WEBHOOK_URL:
        return

    def _threaded_send():
        try:
            payload = {
                "username": "Easy Match Tracker",
                "embeds": [{
                    "title": f"Software {action}",
                    "color": 3066993, # Green
                    "fields": [
                        {"name": "Machine", "value": platform.node(), "inline": True},
                        {"name": "HWID", "value": get_hwid(), "inline": True},
                        {"name": "License", "value": license_info.get("type", "unknown"), "inline": True},
                        {"name": "Expiry", "value": license_info.get("expiry", "unknown"), "inline": True},
                        {"name": "OS", "value": platform.system(), "inline": True}
                    ],
                    "footer": {"text": f"v4.8.1 | {platform.platform()}"}
                }]
            }
            requests.post(WEBHOOK_URL, json=payload, timeout=5)
        except:
            pass # Silent failure

    threading.Thread(target=_threaded_send, daemon=True).start()

def check_remote_block(license_key=None):
    """
    Checks if the current HWID or License Key is blacklisted remotely.
    Returns (is_blocked, message)
    """
    from commercial_config import BLACKLIST_URL
    if not BLACKLIST_URL or "githubusercontent.com/username" in BLACKLIST_URL:
        return False, ""

    try:
        # 1. Fetch Blacklist
        response = requests.get(BLACKLIST_URL, timeout=5)
        if response.status_code != 200:
            return False, ""
        
        blacklist = response.json()
        if not isinstance(blacklist, list):
            return False, ""

        # 2. Check Match
        current_hwid = get_hwid()
        if current_hwid in blacklist:
            return True, f"차단된 기기입니다. (ID: {current_hwid})\n판매자에게 문의하세요."
        
        if license_key and license_key in blacklist:
            return True, f"정지된 라이선스 키입니다. (Key: {license_key})\n새로운 키를 구매해 주세요."

        return False, ""
    except:
        # Fallback to allowed if network is down or URL is invalid 
        # (Safer to let them run once than to block everyone if author's server is down)
        return False, ""
