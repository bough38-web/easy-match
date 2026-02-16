from __future__ import annotations

import os
import datetime
from typing import List, Optional, Callable, Tuple, Dict

import pandas as pd

from utils import norm, smart_format, get_fuzzy_mapper, RAPIDFUZZ_AVAILABLE
from excel_io import read_table_file, write_xlsx
from open_excel import read_table_open, write_to_open_excel

Progress = Optional[Callable[[str, Optional[int]], None]]



# Debug Logger
def _debug_log(msg):
    try:
        log_path = os.path.join(os.path.expanduser("~"), "Desktop", "EasyMatch_Log.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now()}] {msg}\n")
    except:
        pass

def _log(progress: Progress, msg: str, value: int = None) -> None:
    _debug_log(f"Progress: {msg} ({value}%)")
    if progress:
        if value is not None:
            progress(msg, value)
        else:
            progress(msg, None)


def _load_df(cfg: Dict, sheet_cols: List[str]) -> pd.DataFrame:
    _debug_log(f"Loading DF: {cfg.get('type')} - {cfg.get('path') or cfg.get('book')}")
    if cfg.get("type") == "file":
        return read_table_file(cfg["path"], cfg["sheet"], cfg["header"], sheet_cols)
    return read_table_open(cfg["book"], cfg["sheet"], cfg["header"], sheet_cols)


def match_universal(
    base_config: Dict,
    target_config: Dict,
    key_cols: List[str],
    take_cols: List[str],
    out_dir: str,
    options: Dict,
    replacement_rules: Dict[str, Dict[str, str]] | None = None,
    progress: Progress = None,
) -> Tuple[str, str]:
    # safety
    if isinstance(key_cols, str):
        key_cols = [key_cols]
    if isinstance(take_cols, str):
        take_cols = [take_cols]

    key_cols = [str(k).strip() for k in key_cols if str(k).strip()]
    take_cols = [str(c).strip() for c in take_cols if str(c).strip() and c not in key_cols]

    if not key_cols:
        raise ValueError("매칭할 키(Key) 컬럼이 없습니다.")
    if not take_cols:
        raise ValueError("가져올 컬럼이 없습니다.")

    needed_target = list(dict.fromkeys(key_cols + take_cols))

    use_fuzzy = bool(options.get("fuzzy", False))
    use_color = bool(options.get("color", False))

    if len(key_cols) > 1 and use_fuzzy:
        _log(progress, "[INFO] 다중 키 매칭 시 오타 보정은 지원되지 않아 자동 해제됩니다.", 5)
        use_fuzzy = False

    _log(progress, "데이터 로드 중...", 10)
    df_t = _load_df(target_config, needed_target)
    df_b = _load_df(base_config, key_cols)

    # license limit (personal)
    lic_type = (options.get("license_type") or "personal").lower()
    if lic_type == "personal":
        from commercial_config import PERSONAL_MAX_ROWS

        if len(df_b) > PERSONAL_MAX_ROWS or len(df_t) > PERSONAL_MAX_ROWS:
            raise Exception(
                f"개인용 라이선스는 {PERSONAL_MAX_ROWS:,}행 이하만 지원합니다. "
                f"(현재 기준:{len(df_b):,} / 대상:{len(df_t):,})"
            )

    rows_max = max(len(df_b), len(df_t))
    
    # Large Data Warning
    if rows_max > 10000:
        _log(progress, f"대용량 데이터({rows_max:,}행) 처리 중... 잠시만 기다려주세요.", 15)
        
    use_fast = rows_max >= 50000  # auto fast mode for big data

    # replacements (target only)
    if replacement_rules:
        _log(progress, "[Processing] 사용자 정의 치환 규칙 적용 중...", 20)
        for col, rules in replacement_rules.items():
            if col in df_t.columns and isinstance(rules, dict):
                df_t[col] = df_t[col].replace(rules)

    # normalize keys
    _log(progress, "데이터 정규화 중...", 30)
    import numpy as np
    for k in key_cols:
        if k in df_b.columns:
            if use_fast:
                s = df_b[k].astype(str).str.strip()
                df_b[k] = np.where(s.str.endswith(".0"), s.str[:-2], s)
            else:
                df_b[k] = df_b[k].apply(norm)
        if k in df_t.columns:
            if use_fast:
                s = df_t[k].astype(str).str.strip()
                df_t[k] = np.where(s.str.endswith(".0"), s.str[:-2], s)
            else:
                df_t[k] = df_t[k].apply(norm)

    # fuzzy (single key only)
    if use_fuzzy and RAPIDFUZZ_AVAILABLE and len(key_cols) == 1:
        _log(progress, "[AI] 오타 보정(AI Fuzzy) 분석 중...", 40)
        k = key_cols[0]
        if k in df_b.columns and k in df_t.columns:
            mapper = get_fuzzy_mapper(df_b[k], df_t[k], threshold=90)
            if mapper:
                _log(progress, f"총 {len(mapper)}건의 유사 키를 발견하여 보정합니다.")
                df_t[k] = df_t[k].map(mapper).fillna(df_t[k])

    # target dup keys
    if set(key_cols).issubset(df_t.columns):
        dup = int(df_t.duplicated(subset=key_cols).sum())
        if dup:
            _log(progress, f"[WARN] 대상 데이터에 중복 키가 {dup:,}건 있어 첫 번째 값으로만 매칭합니다.")
        df_t = df_t.drop_duplicates(subset=key_cols, keep="first")

    _log(progress, f"매칭 수행 중... (키: {', '.join(key_cols)})", 50)

    if use_fast:
        _log(progress, "[Fast] 대용량 고속 매칭 모드 적용...", 55)

        sep = "||"
        df_b = df_b.copy()
        df_t = df_t.copy()

        # preserve original order
        df_b["_idx"] = df_b.index
        df_b["_key"] = df_b[key_cols].astype(str).agg(sep.join, axis=1)
        df_t["_key"] = df_t[key_cols].astype(str).agg(sep.join, axis=1)

        # one-to-one for mapping
        df_t = df_t.drop_duplicates(subset="_key", keep="first")

        # build mapping series per take col (fast)
        mapping = {col: pd.Series(df_t[col].values, index=df_t["_key"]) for col in take_cols}

        res = pd.DataFrame(index=df_b.index)
        
        # Progress for columns
        total_cols = len(take_cols)
        for i, col in enumerate(take_cols):
            # 60% to 90%
            prog_val = 60 + int((i / total_cols) * 30)
            _log(progress, f"데이터 매칭 생성 중... ({col})", prog_val)
            res[col] = df_b["_key"].map(mapping[col]).fillna("")

        _log(progress, "결과 병합 중...", 90)
        joined = pd.concat([df_b[key_cols], res], axis=1)
        joined = joined.loc[df_b.sort_values("_idx").index]
        joined = joined.drop(columns=[], errors="ignore")
    else:
        joined = pd.merge(df_b.reset_index(), df_t, on=key_cols, how="left")
        if "index" in joined.columns:
            joined = joined.set_index("index").sort_index()
        _log(progress, "매칭 완료, 데이터 정리 중...", 90)

    # select / fill
    final_cols = key_cols + take_cols
    for c in final_cols:
        if c not in joined.columns:
            joined[c] = ""
    joined = joined[final_cols].fillna("")

    # formatting (take cols only)
    for c in take_cols:
        joined[c] = joined[c].map(smart_format)

    # Sanitize entire dataframe to prevent openpyxl crashes (illegal chars)
    from utils import remove_illegal_chars
    _log(progress, "데이터 저장 준비 중 (특수문자 제거)...", 95)
    _debug_log("Sanitizing data...")
    # Apply to all string columns
    for col in joined.select_dtypes(include=['object']).columns:
        joined[col] = joined[col].map(remove_illegal_chars)

    total = len(joined)
    if total:
        # vectorized matched count: any non-empty in take_cols
        import numpy as np
        # Check if any take col has length > 0
        mask = pd.DataFrame(False, index=joined.index, columns=['match'])
        for c in take_cols:
             mask['match'] |= (joined[c].astype(str).str.len() > 0)
        matched = int(mask['match'].sum())
    else:
        matched = 0
    
    _debug_log(f"Matched: {matched}/{total}")
    rate = (matched / total * 100.0) if total else 0.0
    summary = f"[SUCCESS] 총 {total:,}건 중 {matched:,}건 매칭 성공 ({rate:.1f}%)\n[FAIL] 실패: {total - matched:,}건"

    os.makedirs(out_dir, exist_ok=True)
    suffix = base_config["path"] if base_config.get("type") == "file" else base_config.get("book", "base")
    safe = os.path.basename(str(suffix)).split(".")[0]
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"result_{safe}_{ts}.xlsx")
    
    _log(progress, f"파일 저장 중: {os.path.basename(out_path)}", 96)
    print(f"[DEBUG] Saving to: {out_path}")
    _debug_log(f"Saving start: {out_path}")

    try:
        # Explicit try with xlsxwriter first (faster, reliable)
        try:
            import xlsxwriter
            _debug_log("Using xlsxwriter engine...")
            with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
                joined.to_excel(writer, sheet_name="matched", index=False)
            _debug_log("xlsxwriter Save Success.")
        except ImportError:
            # Fallback to default (likely openpyxl)
            _log(progress, "xlsxwriter 없음, 기본 엔진 사용...", 97)
            _debug_log("xlsxwriter module not found. Using default.")
            joined.to_excel(out_path, sheet_name="matched", index=False)
        except Exception as e:
            # xlsxwriter failed? try openpyxl
            print(f"[WARN] xlsxwriter failed: {e}")
            _debug_log(f"xlsxwriter Failed: {e}. Retrying with default...")
            _log(progress, "기본 엔진으로 재시도...", 98)
            joined.to_excel(out_path, sheet_name="matched", index=False)
        
        _debug_log("Final Save Logic Completed.")

        # Extra: Save as CSV blindly if it's too large for some editors to handle XLSX
        if total > 500000:
             csv_path = out_path.replace(".xlsx", ".csv")
             _log(progress, f"대용용 CSV 추가 저장 중: {os.path.basename(csv_path)}", 99)
             _debug_log(f"Saving extra CSV for compatibility: {csv_path}")
             joined.to_csv(csv_path, index=False, encoding="utf-8-sig")
             _debug_log("CSV Save Success.")

    except PermissionError:
        _debug_log("PermissionError encountered.")
        raise Exception(f"저장 실패: 파일이 열려있습니다.\n'{os.path.basename(out_path)}'를 닫아주세요.")
    except Exception as e:
        import traceback
        traceback.print_exc()
        _debug_log(f"Save Exception: {e}")
        raise Exception(f"파일 저장 중 오류 발생: {e}")

    if base_config.get("type") == "open":
        # ... (same as before) ...
        pass # Simplified for clarity, original logic preserved below if needed

    # Original open excel check logic, wrapped safely
    if base_config.get("type") == "open":
        try:
            _log(progress, "엑셀 시트에 결과 입력 중...")
            import pythoncom
            pythoncom.CoInitialize()
            write_to_open_excel(
                base_config["book"],
                base_config["sheet"],
                base_config["header"],
                joined,
                take_cols,
                key_cols,
                use_color,
            )
            _log(progress, "입력 완료.")
        except Exception as e:
            _log(progress, f"[경고] 시트 입력 실패 (파일로만 저장됨): {e}")

    # Final summary with preview data
    preview = joined.head(5) if len(joined) > 0 else None
    
    return out_path, summary, preview
