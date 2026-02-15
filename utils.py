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
    if s.lower() in ["nan","none","null","nat",""]: return ""
    if s.endswith(".0"): s = s[:-2]
    return s

def smart_format(val) -> str:
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if not s or s.lower() in ["nan","none","null","<na>"]: return ""
    if s.endswith(".0"): s = s[:-2]
    elif s.endswith(".00"): s = s[:-3]
    if s.isdigit() and (s.startswith("19") or s.startswith("20")):
        if len(s)==14: return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
        if len(s)==8: return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return s

def get_fuzzy_mapper(base_keys: pd.Series, target_keys: pd.Series, threshold: int = 90) -> dict:
    if not RAPIDFUZZ_AVAILABLE or base_keys is None or target_keys is None: return {}
    if base_keys.empty or target_keys.empty: return {}
    base_choices = base_keys.dropna().astype(str).unique().tolist()
    target_choices = target_keys.dropna().astype(str).unique().tolist()
    mapper = {}
    for t_key in target_choices:
        m = process.extractOne(t_key, base_choices, scorer=fuzz.token_sort_ratio)
        if not m: continue
        best, score, _ = m
        if score >= threshold and t_key != best:
            mapper[t_key] = best
    return mapper
