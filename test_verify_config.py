import os
import sys
from config import APP_DATA_DIR, PRESET_FILE, REPLACE_FILE, LICENSE_FILE, get_system_font

print(f"APP_DATA_DIR: {APP_DATA_DIR}")
print(f"PRESET_FILE: {PRESET_FILE}")
print(f"REPLACE_FILE: {REPLACE_FILE}")
print(f"LICENSE_FILE: {LICENSE_FILE}")
print(f"System Font: {get_system_font()}")

if not os.path.exists(APP_DATA_DIR):
    print("Error: APP_DATA_DIR does not exist.")
    sys.exit(1)

if not os.access(APP_DATA_DIR, os.W_OK):
    print("Error: APP_DATA_DIR is not writable.")
    sys.exit(1)

print("Verification Successful: Config paths are valid and writable.")
