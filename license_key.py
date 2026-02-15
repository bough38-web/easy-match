import hmac
import hashlib
import base64
import json
import datetime
import struct

# In a real production environment, this key should be obscurely hidden or compiled.
# For this Python distribution, we use a hardcoded secret.
# WARNING: Do not share this key if you distribute the source code openly.
SECRET_KEY = b"ExcelMatcher_Secret_Key_2026_X9"

def generate_key(expiry_str: str, license_type: str = "personal") -> str:
    """
    Generates a secure product key.
    Format: EM<VERSION>-<DATA_B32>-<SIGNATURE_B32>
    
    Data payload (JSON): {"e": "2026-12-31", "t": "enterprise"}
    """
    # 1. Prepare Payload
    payload = {
        "e": expiry_str,       # expiry
        "t": license_type[0]   # 'p' or 'e' (save space)
    }
    payload_bytes = json.dumps(payload, separators=(',', ':')).encode('utf-8')
    
    # 2. Sign
    signature = hmac.new(SECRET_KEY, payload_bytes, hashlib.sha256).digest()
    
    # 3. Encode
    # We use Base32 to avoid confusing characters (Verified safe for URL/Text)
    # Remove padding '=' for cleaner look
    data_b32 = base64.b32encode(payload_bytes).decode('utf-8').rstrip('=')
    sig_b32 = base64.b32encode(signature[:10]).decode('utf-8').rstrip('=') # Truncate sig to 10 bytes for shortness
    
    return f"EM-{data_b32}-{sig_b32}"

def validate_key(key_string: str) -> tuple[bool, dict]:
    """
    Validates a product key.
    Returns: (is_valid, data_dict)
    data_dict example: {"expiry": "2026-12-31", "type": "personal"}
    """
    try:
        if not key_string.startswith("EM-"):
            return False, {}
            
        parts = key_string.split('-')
        if len(parts) != 3:
            return False, {}
            
        data_b32 = parts[1]
        sig_b32 = parts[2]
        
        # 1. Restore Padding & Decode
        def add_padding(s):
            return s + '=' * (-len(s) % 8)
            
        payload_bytes = base64.b32decode(add_padding(data_b32))
        
        # 2. Verify Signature
        expected_sig = hmac.new(SECRET_KEY, payload_bytes, hashlib.sha256).digest()
        actual_sig_bytes = base64.b32decode(add_padding(sig_b32))
        
        # We truncated signature to 10 bytes in generation
        if not hmac.compare_digest(expected_sig[:10], actual_sig_bytes):
            return False, {}
            
        # 3. Parse Data
        payload = json.loads(payload_bytes.decode('utf-8'))
        
        expiry = payload.get("e")
        l_type_char = payload.get("t")
        l_type = "enterprise" if l_type_char == 'e' else "personal"
        
        # 4. Check Expiry Logic (Optional here, but good to return data)
        # validation of date vs today is done by caller usually, but we return data.
        
        return True, {"expiry": expiry, "type": l_type}
        
    except Exception:
        return False, {}

if __name__ == "__main__":
    # Test
    today = datetime.date.today().strftime("%Y-%m-%d")
    print(f"Generating key for {today}...")
    key = generate_key(today, "enterprise")
    print(f"Key: {key}")
    
    valid, info = validate_key(key)
    print(f"Validate result: {valid}, {info}")
    
    # Tamper test
    tampered_key = key[:-1] + ('A' if key[-1] != 'A' else 'B')
    print(f"Tampered Key: {tampered_key}")
    print(f"Validate Tampered: {validate_key(tampered_key)}")
