# open_excel.py
from __future__ import annotations

import sys
from typing import List

# xlwings 가용성 플래그 (ui.py에서 import 할 수 있게 "정식"으로 제공)
try:
    import xlwings as xw  # type: ignore
    XLWINGS_AVAILABLE = True
except Exception:
    xw = None
    XLWINGS_AVAILABLE = False


def xlwings_available() -> bool:
    """
    런타임에서 xlwings + Excel 연동 가능 여부를 보수적으로 판단.
    - macOS: Automation 권한/Excel 상태 따라 불안정할 수 있어 예외 대비
    """
    if not XLWINGS_AVAILABLE:
        return False
    try:
        # 간단한 접근 테스트
        _ = xw.apps  # type: ignore
        return True
    except Exception:
        return False


def list_open_books() -> List[str]:
    """
    현재 열려있는 Excel Workbook 이름 목록 반환
    """
    if not xlwings_available():
        return []
    try:
        app = xw.apps.active  # type: ignore
        if app is None:
            return []
        return [wb.name for wb in app.books]
    except Exception:
        return []


def _get_book_by_name(book_name: str):
    if not xlwings_available():
        raise RuntimeError("xlwings/Excel 연동 불가")
    app = xw.apps.active  # type: ignore
    if app is None:
        raise RuntimeError("활성 Excel 앱이 없습니다. Excel을 실행하고 파일을 열어주세요.")
    for wb in app.books:
        if wb.name == book_name:
            return wb
    # 이름이 정확히 일치하지 않을 수도 있어, 부분 매칭 시도(보수적)
    for wb in app.books:
        if book_name in wb.name:
            return wb
    raise RuntimeError(f"열려있는 통합문서에서 '{book_name}' 을(를) 찾지 못했습니다.")


def list_sheets(book_name: str) -> List[str]:
    """
    특정 workbook의 시트 목록 반환
    """
    wb = _get_book_by_name(book_name)
    try:
        return [sh.name for sh in wb.sheets]
    except Exception as e:
        raise RuntimeError(f"시트 목록 조회 실패: {e}") from e


def read_header_open(book_name: str, sheet_name: str, header_row: int = 1) -> List[str]:
    """
    열려있는 Excel에서 header_row 행의 헤더(컬럼명) 목록을 읽어옴.
    - 빈 값은 제거
    """
    wb = _get_book_by_name(book_name)
    try:
        sh = wb.sheets[sheet_name]
    except Exception as e:
        raise RuntimeError(f"시트를 찾지 못했습니다: {sheet_name}") from e

    try:
        # A열부터 우측으로 연속된 헤더를 1행 읽는 방식
        # used_range 기반으로 폭을 잡아도 되지만, 보수적으로 current_region 사용
        rng = sh.range((header_row, 1)).expand("right")
        vals = rng.value

        if vals is None:
            return []
        if isinstance(vals, list):
            # 단일 행이면 list로 나옴
            headers = vals
        else:
            headers = [vals]

        out: List[str] = []
        for v in headers:
            if v is None:
                continue
            s = str(v).strip()
            if s:
                out.append(s)
        return out
    except Exception as e:
        raise RuntimeError(f"헤더 읽기 실패: {e}") from e


def read_table_open(book_name: str, sheet_name: str, header_row: int, usecols: List[str]) -> "pd.DataFrame":
    """
    열려있는 엑셀 시트에서 데이터를 읽어 DataFrame으로 반환
    """
    import pandas as pd
    
    wb = _get_book_by_name(book_name)
    try:
        sh = wb.sheets[sheet_name]
    except Exception as e:
        raise RuntimeError(f"시트 접근 실패: {e}")

    try:
        # 헤더 영역
        rng_header = sh.range((header_row, 1)).expand('right')
        headers = rng_header.value
        if not headers: 
            return pd.DataFrame()
        
        if not isinstance(headers, list):
            headers = [headers]
            
        headers = [str(h).strip() for h in headers]
        
        # 데이터 영역 (헤더 다음 행부터)
        last_row = sh.used_range.last_cell.row
        if last_row <= header_row:
            return pd.DataFrame(columns=[c for c in usecols if c in headers])

        # 전체 데이터 읽기
        data_rng = sh.range((header_row + 1, 1), (last_row, len(headers)))
        data_vals = data_rng.value
        
        if last_row == header_row + 1:
            if isinstance(data_vals, list):
                if len(headers) == 1:
                     data_vals = [[data_vals]]
                else:
                     data_vals = [data_vals]
            else:
                 data_vals = [[data_vals]]
        
        df = pd.DataFrame(data_vals, columns=headers)
        
        # 필요한 컬럼만 필터링
        existing = [c for c in usecols if c in df.columns]
        missing = [c for c in usecols if c not in df.columns]
        
        df = df[existing] if existing else pd.DataFrame()
        for c in missing: 
            df[c] = ""
            
        df = df.reindex(columns=usecols, fill_value="")
        df = df.astype(str).replace(['nan', 'NaN', 'None', '<NA>', 'None'], '')
        
        return df
        
    except Exception as e:
        raise RuntimeError(f"데이터 읽기 실패: {e}")


def write_to_open_excel(book_name: str, sheet_name: str, header_row: int, 
                        df_result: "pd.DataFrame", take_cols: List[str], key_cols: List[str],
                        use_color: bool = True):
    """
    매칭 결과를 다시 엑셀에 기입 (xlwings)
    """
    wb = _get_book_by_name(book_name)
    sh = wb.sheets[sheet_name]
    
    rng_header = sh.range((header_row, 1)).expand('right')
    headers = rng_header.value
    if not isinstance(headers, list): headers = [headers]
    headers = [str(h).strip() for h in headers]
    
    col_map = {h: i+1 for i, h in enumerate(headers)}
    start_row = header_row + 1
    
    for col_name in take_cols:
        col_idx = col_map.get(col_name)
        if not col_idx:
            continue 
            
        col_data = df_result[col_name].values.tolist()
        col_data_vert = [[v] for v in col_data]
        
        start_cell = sh.range(start_row, col_idx)
        target_rng = sh.range(start_cell, (start_row + len(col_data) - 1, col_idx))
        target_rng.value = col_data_vert
        
        if use_color:
            target_rng.color = (255, 255, 204) # Light Yellow