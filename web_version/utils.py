from __future__ import annotations
import pandas as pd
try:
    from rapidfuzz import process, fuzz
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    RAPIDFUZZ_AVAILABLE = False

def norm(s) -> str:
    if s is None: return ""
    s = str(s).strip()
    # Handle pandas 'nan' strings and other empty variants
    if s.lower() in ["nan","none","null","nat",""]: return ""
    
    # Handle dates: "2023-01-01 00:00:00" -> "2023-01-01"
    if len(s) == 19 and s[10:] == " 00:00:00":
        s = s[:10]
        
    # Standardize numeric strings: 
    # Remove commas first (e.g. "1,234" -> "1234")
    if "," in s:
        temp = s.replace(",", "")
        # Only remove if it's truly a number
        try:
            float(temp)
            s = temp
        except: pass
        
    if s.endswith(".0"): 
        s = s[:-2]

    # CASE INSENSITIVE matching is generally preferred in Excel tools
    return s.lower()

def smart_format(val, col_name=None) -> str:
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if not s or s.lower() in ["nan","none","null","<na>"]: return ""
    
    # Generic .0 removal (e.g. "23400.0" -> "23400")
    if s.endswith(".0"): s = s[:-2]
    elif s.endswith(".00"): s = s[:-3]
    
    cn = str(col_name) if col_name else ""
    
    # 1. Monthly Fee (월정료) -> Thousands comma
    if "월정료" in cn:
        try:
            # Remove any existing commas just in case, then format
            f_val = float(s.replace(",", ""))
            return "{:,.0f}".format(f_val)
        except: pass

    # 2. Date Formatting (yyyy-mm-dd)
    # Target columns based on user request + common patterns
    date_cols = ["계약시작일", "계약종료일", "해지일자", "정지시작일자", "정지종료희망일"]
    is_date_col = cn in date_cols or any(k in cn for k in ["시작일", "종료일", "해지일", "일자"])
    
    if is_date_col:
        # Case A: 8-digit string "20230101"
        if s.isdigit() and len(s) == 8:
            return f"{s[:4]}-{s[4:6]}-{s[6:]}"
        # Case B: Timestamp with time "2023-01-01 00:00:00" -> "2023-01-01"
        if len(s) >= 10:
            import re
            m = re.match(r"(\d{4})[-/.]?(\d{2})[-/.]?(\d{2})", s)
            if m:
                return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # 3. Default legacy logic for years/timestamps
    if s.isdigit() and (s.startswith("19") or s.startswith("20")):
        if len(s)==14: return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
        if len(s)==8: return f"{s[:4]}-{s[4:6]}-{s[6:]}"
        
    return s

def get_fuzzy_mapper(base_keys: pd.Series, target_keys: pd.Series, threshold: int = 90, progress_callback=None) -> dict:
    if not RAPIDFUZZ_AVAILABLE or base_keys is None or target_keys is None: return {}
    if base_keys.empty or target_keys.empty: return {}
    base_choices = base_keys.dropna().astype(str).unique().tolist()
    target_choices = target_keys.dropna().astype(str).unique().tolist()
    
    # Breakthrough: Only fuzzy-match items that DON'T have an exact match
    base_set = set(base_choices)
    target_choices = [t for t in target_choices if t and t not in base_set]
    
    # SAFETY: Avoid N*M explosion (Limit to 50M comparisons for responsiveness)
    if not target_choices or (len(base_choices) * len(target_choices) > 50000000):
         return {} 

    mapper = {}
    total = len(target_choices)
    for i, t_key in enumerate(target_choices):
        if progress_callback and i % 100 == 0:
            progress_callback(i, total)
            
        m = process.extractOne(t_key, base_choices, scorer=fuzz.token_sort_ratio)
        if not m: continue
        best, score, _ = m
        if score >= threshold and t_key != best:
            mapper[t_key] = best
    return mapper

import re
def remove_illegal_chars(val):
    """Remove characters that are illegal in Excel (openpyxl)"""
    if not isinstance(val, str):
        return val
    return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', val)

def vectorize_norm(series: pd.Series) -> pd.Series:
    """Vectorized version of norm() for high performance on large datasets."""
    s = series.astype(str).str.strip()
    # Normalize nulls
    low = s.str.lower()
    mask_null = low.isin(["nan", "none", "null", "nat", "", "empty", "<na>"])
    s[mask_null] = ""
    
    # Remove time if it's midnight
    s = s.str.replace(r" 00:00:00$", "", regex=True)
    
    # Standardize numeric strings (remove .0 and commas)
    # and map to lowercase
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"\.0$", "", regex=True)
    
    return s.str.lower()

def vectorize_smart_format(series: pd.Series, col_name: str) -> pd.Series:
    """Vectorized version of smart_format() for high performance."""
    s = series.astype(str).str.strip()
    low = s.str.lower()
    mask_null = low.isin(["nan", "none", "null", "nat", "", "<na>", "undefined"])
    s[mask_null] = ""
    
    # Generic .0 removal
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.str.replace(r"\.00$", "", regex=True)
    
    cn = str(col_name)
    
    # 1. Monthly Fee (월정료) -> Thousands comma
    if "월정료" in cn:
        # Convert to numeric first (vectorized)
        nums = pd.to_numeric(s.str.replace(",", "", regex=False), errors='coerce')
        mask_valid = nums.notna()
        # Still need a way to format as string with commas. 
        # For 1M rows, format() in a map is unavoidable for the final string representation,
        # but we can minimize the work by only applying it to valid numbers.
        s[mask_valid] = nums[mask_valid].apply(lambda x: "{:,.0f}".format(x))
        return s

    # 2. Date Formatting (yyyy-mm-dd)
    date_cols = ["계약시작일", "계약종료일", "해지일자", "정지시작일자", "정지종료희망일"]
    is_date_col = cn in date_cols or any(k in cn for k in ["시작일", "종료일", "해지일", "일자"])
    
    if is_date_col:
        # 8-digit "20230101" -> "2023-01-01"
        mask_8d = s.str.match(r"^\d{8}$")
        s[mask_8d] = s[mask_8d].str.slice(0, 4) + "-" + s[mask_8d].str.slice(4, 6) + "-" + s[mask_8d].str.slice(6, 8)
        
        # Timestamp "2023-01-01 12:34:56" -> "2023-01-01"
        s = s.str.replace(r"^(\d{4})[-/.]?(\d{2})[-/.]?(\d{2}).*$", r"\1-\2-\3", regex=True)
        
    return s

def apply_expert_norm(series: pd.Series) -> pd.Series:
    """Breakthrough: Normalize unique values only. 100x faster for 1M rows with repeating data."""
    if series.empty: return series
    u = series.unique()
    u_norm = vectorize_norm(pd.Series(u))
    mapping = dict(zip(u, u_norm))
    return series.map(mapping)

def apply_expert_format(series: pd.Series, col_name: str) -> pd.Series:
    """Breakthrough: Format unique values only. High performance for millions of rows."""
    if series.empty: return series
    if series.dtype != 'object' and "월정료" not in str(col_name):
        return series.astype(str)
    u = series.unique()
    u_fmt = vectorize_smart_format(pd.Series(u), col_name)
    mapping = dict(zip(u, u_fmt))
    return series.map(mapping)
