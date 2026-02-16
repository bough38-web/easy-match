import pandas as pd
import os
from matcher import match_universal

def run_test():
    print("=== Deep Verification Test ===")
    
    # 1. Create Data with Edge Cases
    # Base: Int, Float-as-int, String-as-int, Whitespace, Empty
    df_b = pd.DataFrame({
        "ID": [101, 102.0, "103", " 104 ", 105.0, "106.0", None, "Nan"],
        "Name": ["A", "B", "C", "D", "E", "F", "G", "H"]
    })
    
    # Target: Mixed types corresponding to Base
    df_t = pd.DataFrame({
        "ID": ["101", 102, 103.0, 104, 105, "106", "107", " "],
        "Value": ["Matched-101", "Matched-102", "Matched-103", "Matched-104", "Matched-105", "Matched-106", "Unmatched", "Unmatched"]
    })
    
    # Save files
    df_b.to_excel("deep_base.xlsx", index=False)
    df_t.to_excel("deep_target.xlsx", index=False)
    
    b_cfg = {"type": "file", "path": os.path.abspath("deep_base.xlsx"), "sheet": 0, "header": 1}
    t_cfg = {"type": "file", "path": os.path.abspath("deep_target.xlsx"), "sheet": 0, "header": 1}
    
    out_dir = "deep_out"
    
    # 2. Run Match (Normal Mode)
    print("\n--- Normal Mode Test ---")
    out, summary = match_universal(
        b_cfg, t_cfg, ["ID"], ["Value"], out_dir, 
        {"fuzzy": False}, None, lambda m, v: print(f"[Normal] {m}")
    )
    print(f"Saved: {out}")
    verify_result(out, "Normal")

    # 3. Fast Mode Test (Simulate by setting extremely low threshold in matcher if possible, 
    # but since I can't change code easily, I will just replicate the logic or rely on normal mode validation first).
    # Wait, I can't force fast mode without 50k rows. 
    # I'll create large partial data to force it.
    
    print("\n--- Fast Mode Test (Forced via Data Size) ---")
    # Duplicate rows to reach 50000 (Limit is >50000 raise, so 50000 is OK)
    # 8 rows * 6250 = 50000
    df_b_large = pd.concat([df_b] * 6250, ignore_index=True)
    df_t_large = pd.concat([df_t] * 6250, ignore_index=True)
    
    df_b_large.to_excel("deep_base_large.xlsx", index=False)
    df_t_large.to_excel("deep_target_large.xlsx", index=False)
    
    b_cfg_l = {"type": "file", "path": os.path.abspath("deep_base_large.xlsx"), "sheet": 0, "header": 1}
    t_cfg_l = {"type": "file", "path": os.path.abspath("deep_target_large.xlsx"), "sheet": 0, "header": 1}
    
    out_l, summary_l = match_universal(
        b_cfg_l, t_cfg_l, ["ID"], ["Value"], out_dir, 
        {"fuzzy": False}, None, lambda m, v: print(f"[Fast] {m} ({v})")
    )
    verify_result_head(out_l, "Fast")


def verify_result(path, mode):
    df = pd.read_excel(path)
    print(f"[{mode}] Result Head:")
    print(df[["ID", "Value"]].head(10))
    
    # Check expected matches
    # 101 (Int) vs "101" (Str) -> Should Match in Normal
    # 102.0 (Float) vs 102 (Int) -> Should Match
    # "103" (Str) vs 103.0 (Float) -> Should Match
    # " 104 " (Space) vs 104 (Int) -> Should Match
    # 105.0 (Float) vs 105 (Int) -> Should Match
    # "106.0" (Str) vs "106" (Str) -> Should Match (if norm handles .0 in string)
    
    matched_count = df[df["Value"].str.contains("Matched", na=False)].shape[0]
    print(f"[{mode}] Matched Count: {matched_count} / {len(df)}")

def verify_result_head(path, mode):
    df = pd.read_excel(path)
    print(f"[{mode}] Result Head (First 8 rows):")
    # We only care about the first 8 original rows pattern
    print(df[["ID", "Value"]].head(8))
    matched_count = df["Value"].head(8).str.contains("Matched", na=False).sum()
    print(f"[{mode}] First 8 Rows Matched: {matched_count}/6 Expected")

if __name__ == "__main__":
    try:
        run_test()
    except Exception as e:
        import traceback
        traceback.print_exc()
