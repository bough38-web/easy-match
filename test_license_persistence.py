import os
import datetime
from license_manager import _check_persistent_trial

print("Testing Persistent Trial...")

# 1. Clear existing
home = os.path.expanduser("~")
trial_file = os.path.join(home, ".excelmatcher_trial")

# Backup real file if exists
backup = None
if os.path.exists(trial_file):
    with open(trial_file, 'r') as f:
        backup = f.read()
    os.remove(trial_file)
    print("Backed up and removed existing trial file.")

try:
    # 2. First Run
    ok, msg = _check_persistent_trial()
    print(f"First Run: {ok}, {msg}")
    if not os.path.exists(trial_file):
        print("FAIL: File not created.")
    else:
        print("PASS: File created.")

    # 3. Modify file to be old (40 days ago)
    old_date = datetime.date.today() - datetime.timedelta(days=40)
    with open(trial_file, "w") as f:
        f.write(old_date.strftime("%Y-%m-%d"))
    print(f"Modified file to: {old_date}")

    # 4. Check again
    ok, msg = _check_persistent_trial()
    print(f"Expired Run: {ok}, {msg}")

    if not ok and "만료" in msg:
        print("PASS: Correctly detected expiration.")
    else:
        print("FAIL: Did not detect expiration.")

finally:
    # 5. Restore
    if backup:
        with open(trial_file, 'w') as f:
            f.write(backup)
        print("Restored original trial file.")
    elif os.path.exists(trial_file):
        os.remove(trial_file)
