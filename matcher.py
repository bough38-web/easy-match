from __future__ import annotations

import os
import datetime
from typing import List, Optional, Callable, Tuple, Dict

import pandas as pd

from utils import norm, smart_format, get_fuzzy_mapper, RAPIDFUZZ_AVAILABLE
from excel_io import read_table_file, write_xlsx
from open_excel import read_table_open, write_to_open_excel

Progress = Optional[Callable[[str], None]]


def _log(progress: Progress, msg: str) -> None:
    if progress:
        progress(msg)


def _load_df(cfg: Dict, sheet_cols: List[str]) -> pd.DataFrame:
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
        _log(progress, "[INFO] 다중 키 매칭 시 오타 보정은 지원되지 않아 자동 해제됩니다.")
        use_fuzzy = False

    _log(progress, "데이터 로드 중...")
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
    use_fast = rows_max >= 50000  # auto fast mode for big data

    # replacements (target only)
    if replacement_rules:
        _log(progress, "[Processing] 사용자 정의 치환 규칙 적용 중...")
        for col, rules in replacement_rules.items():
            if col in df_t.columns and isinstance(rules, dict):
                df_t[col] = df_t[col].replace(rules)

    # normalize keys
    _log(progress, "데이터 정규화 중...")
    for k in key_cols:
        if k in df_b.columns:
            df_b[k] = df_b[k].astype(str).str.strip() if use_fast else df_b[k].apply(norm)
        if k in df_t.columns:
            df_t[k] = df_t[k].astype(str).str.strip() if use_fast else df_t[k].apply(norm)

    # fuzzy (single key only)
    if use_fuzzy and RAPIDFUZZ_AVAILABLE and len(key_cols) == 1:
        _log(progress, "[AI] 오타 보정(AI Fuzzy) 분석 중...")
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

    _log(progress, f"매칭 수행 중... (키: {', '.join(key_cols)})")

    if use_fast:
        _log(progress, "[Fast] 대용량 고속 매칭 모드 적용...")

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
        for col in take_cols:
            res[col] = df_b["_key"].map(mapping[col]).fillna("")

        joined = pd.concat([df_b[key_cols], res], axis=1)
        joined = joined.loc[df_b.sort_values("_idx").index]
        joined = joined.drop(columns=[], errors="ignore")
    else:
        joined = pd.merge(df_b.reset_index(), df_t, on=key_cols, how="left")
        if "index" in joined.columns:
            joined = joined.set_index("index").sort_index()

    # select / fill
    final_cols = key_cols + take_cols
    for c in final_cols:
        if c not in joined.columns:
            joined[c] = ""
    joined = joined[final_cols].fillna("")

    # formatting (take cols only)
    for c in take_cols:
        joined[c] = joined[c].map(smart_format)

    total = len(joined)
    if total:
        # vectorized matched count: any non-empty in take_cols
        arr = joined[take_cols].astype(str).to_numpy()
        matched = int((pd.DataFrame(arr).applymap(lambda x: str(x).strip() != "").any(axis=1)).sum())
    else:
        matched = 0
    rate = (matched / total * 100.0) if total else 0.0
    summary = f"[SUCCESS] 총 {total:,}건 중 {matched:,}건 매칭 성공 ({rate:.1f}%)\n[FAIL] 실패: {total - matched:,}건"

    os.makedirs(out_dir, exist_ok=True)
    suffix = base_config["path"] if base_config.get("type") == "file" else base_config.get("book", "base")
    safe = os.path.basename(str(suffix)).split(".")[0]
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"result_{safe}_{ts}.xlsx")

    write_xlsx(out_path, joined, sheet_name="matched")

    if base_config.get("type") == "open":
        _log(progress, "엑셀 시트에 결과 입력 중...")
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

    return out_path, summary
