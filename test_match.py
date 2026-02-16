from matcher import match_universal
import os

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
out_dir = os.path.abspath("test_output")
options = {"fuzzy": False, "color": False}
replace_rules = {}

def progress_cb(msg, val=None):
    print(f"[PROGRESS] {msg} ({val})")

try:
    print("Starting match...")
    out, summary = match_universal(b_cfg, t_cfg, keys, take, out_dir, options, replace_rules, progress_cb)
    print("Result:", out)
    print("Summary:", summary)
except Exception as e:
    import traceback
    traceback.print_exc()
