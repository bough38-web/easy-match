from __future__ import annotations

import os
import datetime
import time
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
    filters: Dict = None,
    progress: Progress = None,
    cancel_check: Callable[[], bool] = lambda: False,
) -> Tuple[str, str, List[dict]]:
    start_time = time.time()

    def log_progress(msg, val=None):
        elapsed = time.time() - start_time
        msg_with_time = f"{msg} ({elapsed:.1f}s)"
        _log(progress, msg_with_time, val)

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
        log_progress("[INFO] 다중 키 매칭 시 오타 보정은 지원되지 않아 자동 해제됩니다.", 5)
        use_fuzzy = False

    log_progress("데이터 로드 중...", 10)
    df_b = _load_df(base_config, key_cols)
    if cancel_check(): raise InterruptedError()
    
    df_t = _load_df(target_config, key_cols + take_cols)  # load keys for matching + takes
    if cancel_check(): raise InterruptedError()
    
    # license limit (personal)
    lic_type = (options.get("license_type") or "personal").lower()
    if lic_type == "personal":
        from commercial_config import PERSONAL_MAX_ROWS

        if len(df_b) > PERSONAL_MAX_ROWS or len(df_t) > PERSONAL_MAX_ROWS:
            from commercial_config import CONTACT_INFO
            raise Exception(
                f"현재 라이선스는 {PERSONAL_MAX_ROWS:,}행 이하만 지원합니다.\n"
                f"(현재 데이터 - 기준: {len(df_b):,} / 대상: {len(df_t):,}행)\n\n"
                f"100만 행 이상의 대용량 데이터 처리는 커스텀 버전이 필요합니다.\n"
                f"문의: {CONTACT_INFO}"
            )

    rows_max = max(len(df_b), len(df_t))
    
    # Large Data Warning
    if rows_max > 10000:
        log_progress(f"대용량 데이터({rows_max:,}행) 처리 중... 잠시만 기다려주세요.", 15)
        
    use_fast = rows_max >= 50000  # auto fast mode for big data

    # replacements (target only)
    if replacement_rules:
        log_progress("[Processing] 사용자 정의 치환 규칙 적용 중...", 20)
        for col, rules in replacement_rules.items():
            if col in df_t.columns and isinstance(rules, dict):
                df_t[col] = df_t[col].replace(rules)

    # Filtering Logic
    if filters:
        log_progress("데이터 필터링 적용 중...", 22)
        
        def apply_multi_filters(df, f_list, label):
            if not f_list: return df
            if isinstance(f_list, dict): f_list = [f_list]
            
            res_df = df.copy()
            for f in f_list:
                if cancel_check(): raise InterruptedError()
                col = f.get("col")
                op = f.get("op", "==")
                val = f.get("keyword") or f.get("value")
                
                if col not in res_df.columns: continue
                if val in ["(값 선택)", "(데이터 없음)", None, ""]: continue
                
                try:
                    # Numeric Conversion if possible
                    if op in [">=", "<=", ">", "<"]:
                        f_val = float(val)
                        col_series = pd.to_numeric(res_df[col], errors='coerce')
                    else:
                        f_val = str(val)
                        col_series = res_df[col].astype(str)
                    
                    if op == "==": 
                        if val == "(값 있음)":
                            res_df = res_df[res_df[col].astype(str).str.strip().replace(['nan','NaN','None',''], None).notnull()]
                        elif val == "(값 없음)":
                            res_df = res_df[res_df[col].astype(str).str.strip().replace(['nan','NaN','None',''], None).isnull()]
                        else:
                            res_df = res_df[col_series == f_val]
                    elif op == ">=": res_df = res_df[col_series >= f_val]
                    elif op == "<=": res_df = res_df[col_series <= f_val]
                    elif op == ">": res_df = res_df[col_series > f_val]
                    elif op == "<": res_df = res_df[col_series < f_val]
                    elif op == "Exist":
                        res_df = res_df[res_df[col].astype(str).str.strip().replace(['nan','NaN','None',''], None).notnull()]
                    elif op == "Not Exist":
                        res_df = res_df[res_df[col].astype(str).str.strip().replace(['nan','NaN','None',''], None).isnull()]
                    
                    _debug_log(f"[Filter] {label} ({col} {op} {val})")
                except Exception as fe:
                    _debug_log(f"[Warning] 필터 적용 실패 ({col}): {fe}")
            return res_df

        # Apply Base Filters
        base_filters = filters.get("base_multi", [])
        if not base_filters and (filters.get("base") or filters.get("base_prefix")):
            base_filters = [filters.get("base") or filters.get("base_prefix")]
        df_b = apply_multi_filters(df_b, base_filters, "기준")

        # Apply Target Filters
        tgt_filters = filters.get("target_multi", [])
        if not tgt_filters and filters.get("target_prefix"):
            tgt_filters = [filters.get("target_prefix")]
        df_t = apply_multi_filters(df_t, tgt_filters, "대상")
        
        # Target Filter: Multiple exact value match (dropdown based / old advanced)
        target_fs = filters.get("target_advanced", [])
        for tf in target_fs:
            if cancel_check(): raise InterruptedError()
            col, vals = tf.get("col"), tf.get("values")
            if col in df_t.columns and vals:
                if "(값 있음)" in vals:
                    df_t = df_t[df_t[col].astype(str).str.strip().replace(['nan','NaN','None',''], None).notnull()].copy()
                elif "(값 없음)" in vals:
                    df_t = df_t[df_t[col].astype(str).str.strip().replace(['nan','NaN','None',''], None).isnull()].copy()
                else:
                    df_t = df_t[df_t[col].astype(str).isin(vals)].copy()
                log_progress(f"[Filter] 대상 데이터(고급): {len(df_t):,}건 (필터: {', '.join(vals)})")

    # Expert Option: Top 10
    if options.get("top10") and not df_t.empty:
        log_progress("[Expert] 상위 10개 데이터만 추출 중...", 25)
        # If there's a numeric column to sort by, we could ask, but usually it's just first 10
        # or we sort by the first column as a proxy for 'relevance' if it's already sorted
        df_t = df_t.head(10).copy()

    if df_b.empty:
        raise ValueError("필터 결과 기준 데이터가 비어 있습니다. 매칭을 진행할 수 없습니다.")
    if df_t.empty:
        raise ValueError("필터 결과 대상 데이터가 비어 있습니다. 매칭을 진행할 수 없습니다.")

    # normalize keys
    log_progress("데이터 정규화 중...", 30)
    
    # 1. Deduplicate Target by keys (we only need one match if multiple?)
    #    Actually current logic: keep all targets? 
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
        log_progress("[AI] 오타 보정(AI Fuzzy) 분석 중...", 40)
        k = key_cols[0]
        if k in df_b.columns and k in df_t.columns:
            mapper = get_fuzzy_mapper(df_b[k], df_t[k], threshold=90)
            if mapper:
                log_progress(f"총 {len(mapper)}건의 유사 키를 발견하여 보정합니다.")
                df_t[k] = df_t[k].map(mapper).fillna(df_t[k])

    # target dup keys
    if set(key_cols).issubset(df_t.columns):
        dup = int(df_t.duplicated(subset=key_cols).sum())
        if dup:
            log_progress(f"[WARN] 대상 데이터에 중복 키가 {dup:,}건 있어 첫 번째 값으로만 매칭합니다.")
        df_t = df_t.drop_duplicates(subset=key_cols, keep="first")

    if not use_fuzzy:
        # FAST MERGE (Exact)
        log_progress(f"매칭 수행 중... (키: {', '.join(key_cols)})", 50)
        
        # We need to temporarily rename columns to avoid collision if base and target have same columns
        # This is handled by only selecting key_cols from df_b and then merging with df_t
        # The final selection `joined = joined[final_cols]` will ensure only desired columns are kept.
        
        if use_fast:
            log_progress("[Fast] 대용량 고속 매칭 모드 적용...", 55)

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
                if cancel_check(): raise InterruptedError()
                # 60% to 90%
                prog_val = 60 + int((i / total_cols) * 30)
                log_progress(f"데이터 매칭 생성 중... ({col})", prog_val)
                res[col] = df_b["_key"].map(mapping[col]).fillna("")

            log_progress("결과 병합 중...", 90)
            joined = pd.concat([df_b[key_cols], res], axis=1)
            joined = joined.loc[df_b.sort_values("_idx").index]
            joined = joined.drop(columns=[], errors="ignore")
        else:
            joined = pd.merge(df_b.reset_index(), df_t, on=key_cols, how="left")
            if "index" in joined.columns:
                joined = joined.set_index("index").sort_index()
            log_progress("매칭 완료, 데이터 정리 중...", 90)
    else: # Fuzzy matching (single key only, already checked)
        k = key_cols[0]
        log_progress(f"Fuzzy 매칭 중... (키: {k})", 50)
        
        # Prepare data
        # Base keys (allow duplicates in base, we iterate unique for speed then map back)
        b_series = df_b[k].astype(str)
        b_uniques = b_series.unique()
        
        # Target keys (deduplicated for lookup)
        # We need a map from t_key -> t_row_data
        df_t = df_t.drop_duplicates(subset=[k])
        t_keys = df_t[k].astype(str).tolist()
        
        # We'll build a mapping: base_val -> target_val
        # using rapidfuzz if available
        key_map = {}
        
        if RAPIDFUZZ_AVAILABLE:
            from rapidfuzz import process, fuzz
            total_u = len(b_uniques)
            
            for i, b_val in enumerate(b_uniques):
                if i % 100 == 0:
                    if cancel_check(): raise InterruptedError()
                    # Progress 50% -> 90%
                    prog = 50 + int((i / total_u) * 40)
                    log_progress(f"Fuzzy 정밀 분석 중... ({i}/{total_u})", prog)
                
                # Check exact first
                if b_val in t_keys:
                    key_map[b_val] = b_val
                    continue
                    
                # Fuzzy check
                # score_cutoff=80 (default in typical utils)
                match = process.extractOne(b_val, t_keys, scorer=fuzz.token_sort_ratio, score_cutoff=80)
                if match:
                    res_key, score, _ = match
                    key_map[b_val] = res_key
        else:
            log_progress("[WARN] rapidfuzz 모듈 없음. 정확한 일치만 수행합니다.")
            # Fallback exact
            exact_set = set(t_keys)
            for b_val in b_uniques:
                if b_val in exact_set:
                    key_map[b_val] = b_val

        # Apply mapping to create a join key
        log_progress("매칭 결과 병합 중...", 90)
        df_b['_join_key'] = df_b[k].astype(str).map(key_map)
        
        # Prevent "ValueError: You are trying to merge on float64 and object columns"
        # If map returns all NaNs (no matches), pandas infers float64. Target key is object (str).
        df_b['_join_key'] = df_b['_join_key'].astype(object)
        
        # Merge
        joined = pd.merge(df_b, df_t, left_on='_join_key', right_on=k, how='left', suffixes=('', '_tgt'))
        
        # Cleanup
        if '_join_key' in joined.columns:
            joined.drop(columns=['_join_key'], inplace=True)
        # Handle key collision in columns (if k is in both, merge might rename)
        # We want to keep base key as primary?
        # Typically we keep Base Key. Target Key is redundant if matched.


    # select / fill
    final_cols = key_cols + take_cols
    for c in final_cols:
        if c not in joined.columns:
            joined[c] = ""
    joined = joined[final_cols].fillna("")

    # formatting (take cols only)
    for c in take_cols:
        joined[c] = joined[c].map(smart_format)

    # Sanitize Column Names (IMPORTANT for software compatibility like Bree/LibreOffice)
    from utils import remove_illegal_chars
    log_progress("파일 헤더 정리 중...", 94)
    joined.columns = [remove_illegal_chars(str(c)) for c in joined.columns]
    # Update expected lists with sanitized names
    sanitized_take = [remove_illegal_chars(str(c)) for c in take_cols]
    sanitized_key = [remove_illegal_chars(str(c)) for c in key_cols]

    # Sanitize entire dataframe to prevent openpyxl crashes (illegal chars)
    log_progress("데이터 저장 준비 중 (특수문자 제거)...", 95)
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

        # Save Condition: Match Only
        if options.get("match_only"):
            before_len = len(joined)
            joined = joined[mask['match']].copy()
            total = len(joined)
            matched = total # All remaining are matched
            log_progress(f"매칭 미성공 데이터 제외 완료 ({before_len} -> {total}건)", 92)

    else:
        matched = 0
    
    _debug_log(f"Matched: {matched}/{total}")
    rate = (matched / total * 100.0) if total else 0.0
    summary = f"[SUCCESS] 총 {total:,}건 중 {matched:,}건 매칭 성공 ({rate:.1f}%)\n[FAIL] 실패: {total - matched:,}건"

    os.makedirs(out_dir, exist_ok=True)
    suffix = base_config["path"] if base_config.get("type") == "file" else base_config.get("book", "base")
    safe = os.path.basename(str(suffix)).split(".")[0]
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Threshold-based format selection: XLSX for small data, CSV for large data
    save_as_csv = total > 50000
    ext = ".csv" if save_as_csv else ".xlsx"
    out_path = os.path.join(out_dir, f"result_{safe}_{ts}{ext}")
    
    log_progress(f"파일 저장 중: {os.path.basename(out_path)}", 96)
    print(f"[DEBUG] Saving to: {out_path}")
    _debug_log(f"Saving start: {out_path}")

    try:
        if save_as_csv:
            log_progress(f"CSV(BOM) 저장 중: {os.path.basename(out_path)}", 96)
            _debug_log(f"Saving as CSV (UTF-8-SIG): {out_path}")
            joined.to_csv(out_path, index=False, encoding="utf-8-sig")
            _debug_log("CSV Save Success.")
        else:
            # Explicit try with xlsxwriter first (faster, reliable)
            try:
                import xlsxwriter
                _debug_log("Using xlsxwriter engine with optimized settings...")
                # strings_to_urls=False prevents crashes on long strings that look like URLs
                with pd.ExcelWriter(out_path, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
                    joined.to_excel(writer, sheet_name="matched", index=False)
                    # Auto-filter and freeze pane for professional look
                    worksheet = writer.sheets['matched']
                    worksheet.freeze_panes(1, 0)
                _debug_log("xlsxwriter Save Success.")
            except ImportError:
                # Fallback to default (likely openpyxl)
                log_progress("xlsxwriter 없음, 기본 엔진 사용...", 97)
                _debug_log("xlsxwriter module not found. Using default.")
                joined.to_excel(out_path, sheet_name="matched", index=False)
            except Exception as e:
                # xlsxwriter failed? try openpyxl
                print(f"[WARN] xlsxwriter failed: {e}")
                _debug_log(f"xlsxwriter Failed: {e}. Retrying with default...")
                log_progress("기본 엔진으로 재시도...", 98)
                joined.to_excel(out_path, sheet_name="matched", index=False)
        
        _debug_log("Final Save Logic Completed.")

    except PermissionError:
        _debug_log("PermissionError encountered.")
        raise Exception(f"저장 실패: 파일이 열려있습니다.\n'{os.path.basename(out_path)}'를 닫아주세요.")
    except Exception as e:
        import traceback
        traceback.print_exc()
        _debug_log(f"Save Exception: {e}")
        raise Exception(f"파일 저장 중 오류 발생: {e}")

    # Process "Open Excel" if applicable
    if base_config.get("type") == "open":
        try:
            log_progress("엑셀 시트에 결과 입력 중...")
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
            log_progress("입력 완료.")
            # Preview for open excel
            if len(df_t) > 10:
                df_t = df_t.head(10)
                log_progress(f"Top 10 추출 완료 ({len(df_t)}건)")
        except Exception as e:
            log_progress(f"[경고] 시트 입력 실패 (파일로만 저장됨): {e}")

    # Prepare preview for UI (must be DataFrame)
    preview = joined.head(5) if len(joined) > 0 else None
    
    return out_path, summary, preview
