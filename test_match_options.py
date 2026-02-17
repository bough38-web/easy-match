from matcher import match_universal
import os
import pandas as pd

b_cfg = {
    "type": "file",
    "path": os.path.abspath("test_data_base.xlsx"),
    "sheet": 0,
    "header": 1
}
t_cfg = {
    "type": "file",
    "path": os.path.abspath("test_data_target.xlsx"),
    "sheet": 0,
    "header": 1
}
keys = ["Key"]
take = ["Salary"]
out_dir = os.path.abspath("test_output_options")
options = {"fuzzy": False, "color": False, "match_only": True}
replace_rules = {}

def progress_cb(msg, val=None):
    print(f"[PROGRESS] {msg} ({val})")

try:
    print("Starting match with match_only=True...")
    filters = {}
    out, summary, preview = match_universal(b_cfg, t_cfg, keys, take, out_dir, options, replace_rules, filters, progress=progress_cb)
    print("Result:", out)
    print("Summary:", summary)
    
    # Verify result
    df = pd.read_excel(out, sheet_name="matched")
    print(f"Result Rows: {len(df)}")
    
    if len(df) == 2:
        print("TEST PASSED: Only matched rows saved.")
    else:
        print(f"TEST FAILED: Expected 2 rows, got {len(df)}")

except Exception as e:
    import traceback
    traceback.print_exc()
