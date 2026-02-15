import os
import sys
import platform

# -----------------------------
# Cross-Platform Path Resolution
# -----------------------------
def get_app_data_dir():
    """
    Returns the application data directory safely for both macOS and Windows.
    e.g. ~/.excelmatcher/
    """
    home = os.path.expanduser("~")
    app_dir = os.path.join(home, ".excelmatcher")
    
    # Create the directory if it doesn't exist
    if not os.path.exists(app_dir):
        try:
            os.makedirs(app_dir)
        except OSError:
            # Fallback to current directory if permissions fail
            return os.getcwd()
            
    return app_dir

# Global constants for paths
# Global constants for paths
APP_DATA_DIR = get_app_data_dir()
PRESET_FILE = os.path.join(APP_DATA_DIR, "presets.json")
REPLACE_FILE = os.path.join(APP_DATA_DIR, "replacements.json")

# Prioritize local license.lic (Portable Mode)
_local_lic = os.path.join(os.getcwd(), "license.lic")
if os.path.exists(_local_lic):
    LICENSE_FILE = _local_lic
else:
    LICENSE_FILE = os.path.join(APP_DATA_DIR, "license.lic")

# -----------------------------
# System Font Detection
# -----------------------------
def get_system_font():
    """
    Returns a suitable UI font based on the operating system.
    """
    system = platform.system()
    if system == "Windows":
        return ("Malgun Gothic", 9)
    elif system == "Darwin":  # macOS
        # 'AppleSDGothicNeo-Regular' caused CoreText crash on some systems.
        # Fallback to a safer standard font or system default.
        return ("Helvetica", 12)
    else:
        return ("TkDefaultFont", 10)
