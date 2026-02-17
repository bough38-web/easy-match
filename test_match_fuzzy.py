from matcher import match_universal
import os
import pandas as pd

# Create dummy data for fuzzy match
# Base: ["Apple", "Banana", "Orange"]
# Target: ["Appl", "Banan", "Orang"] -> Should match
b_data = {"Key": ["Apple", "Banana", "Orange"], "Val": [1, 2, 3]}
t_data = {"Key": ["Appl", "Banan", "Orang"], "Price": [100, 200, 300]}

df_b = pd.DataFrame(b_data)
df_t = pd.DataFrame(t_data)

df_b.to_excel("test_fuzzy_base.xlsx", index=False)
df_t.to_excel("test_fuzzy_target.xlsx", index=False)

b_cfg = {
    "type": "file",
    "path": os.path.abspath("test_fuzzy_base.xlsx"),
    "sheet": 0,
    "header": 1 # 1-based index means row 0
}
# Actually in match_universal it expects 0-based index?
# Let's check excel_io.
# ui.py passes index.
# Let's assume 0 for row 0.

b_cfg = {
    "type": "file",
    "path": os.path.abspath("test_fuzzy_base.xlsx"),
    "sheet": 0,
    "header": 1
}
t_cfg = {
    "type": "file",
    "path": os.path.abspath("test_fuzzy_target.xlsx"),
    "sheet": 0,
    "header": 1
}

keys = ["Key"]
take = ["Price"]
out_dir = os.path.abspath("test_output_fuzzy")
options = {"fuzzy": True, "color": False, "match_only": False}
replace_rules = {}

def progress_cb(msg, val=None):
    print(f"[PROGRESS] {msg} ({val})")

try:
    print("Starting match with Fuzzy=True...")
    filters = {}
    out, summary, preview = match_universal(b_cfg, t_cfg, keys, take, out_dir, options, replace_rules, filters, progress=progress_cb)
    print("Result:", out)
    print("Summary:", summary)
    
    # Verify result
    df = pd.read_excel(out)
    print("Result Data:")
    print(df)
    
    # Check if Price is filled
    if df["Price"].notna().all():
         print("TEST PASSED: Fuzzy match successful.")
    else:
         print("TEST FAILED: Some rows not matched.")

except Exception as e:
    import traceback
    traceback.print_exc()
