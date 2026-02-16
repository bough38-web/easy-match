import hmac
import hashlib
import base64
import json
import datetime

SECRET_KEY = b"ExcelMatcher_Secret_Key_2026_X9"

def generate_key(expiry_str, license_type="personal"):
    payload = {
        "e": expiry_str,
        "t": license_type[0]
    }
    payload_bytes = json.dumps(payload, separators=(',', ':')).encode('utf-8')
    signature = hmac.new(SECRET_KEY, payload_bytes, hashlib.sha256).digest()
    data_b32 = base64.b32encode(payload_bytes).decode('utf-8').rstrip('=')
    sig_b32 = base64.b32encode(signature[:10]).decode('utf-8').rstrip('=')
    return f"EM-{data_b32}-{sig_b32}"

if __name__ == "__main__":
    # Generate Enterprise key for test
    key = generate_key("2026-12-31", "enterprise")
    print(f"ENTERPRISE_KEY: {key}")
