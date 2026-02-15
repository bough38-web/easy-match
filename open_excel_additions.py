import pandas as pd

def read_table_open(book_name: str, sheet_name: str, header_row: int, usecols: List[str]) -> pd.DataFrame:
    """
    열려있는 엑셀 시트에서 데이터를 읽어 DataFrame으로 반환
    """
    import pandas as pd
    
    wb = _get_book_by_name(book_name)
    try:
        sh = wb.sheets[sheet_name]
    except Exception as e:
        raise RuntimeError(f"시트 접근 실패: {e}")

    # 헤더 읽어서 컬럼 위치 파악
    # (xlwings range expand 사용)
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
        # expand('table')을 쓰면 끊긴 데이터에서 멈출 수 있으므로
        # used_range를 쓰거나, 아니면 헤더 폭 만큼 아래로 읽기
        
        # 여기서는 간단히: UsedRange의 마지막 행 확인
        last_row = sh.used_range.last_cell.row
        if last_row <= header_row:
            return pd.DataFrame(columns=[c for c in usecols if c in headers])

        # 전체 데이터 읽기 (속도를 위해 한번에 값만 가져옴)
        # 범위: (header_row+1, 1) ~ (last_row, len(headers))
        data_rng = sh.range((header_row + 1, 1), (last_row, len(headers)))
        data_vals = data_rng.value
        
        # 데이터가 1행인 경우 리스트의 리스트가 아닐 수 있음 처리
        if last_row == header_row + 1:
            if isinstance(data_vals, list):
                # 컬럼이 여러 개면 list, 1개면 값
                if len(headers) == 1:
                     data_vals = [[data_vals]]
                else:
                     data_vals = [data_vals]
            else:
                # 값 자체가 하나
                 data_vals = [[data_vals]]
        
        df = pd.DataFrame(data_vals, columns=headers)
        
        # 필요한 컬럼만 필터링 + 없는 컬럼 빈값 추가
        existing = [c for c in usecols if c in df.columns]
        missing = [c for c in usecols if c not in df.columns]
        
        df = df[existing] if existing else pd.DataFrame()
        for c in missing: 
            df[c] = ""
            
        # 순서 정렬 및 Nan 처리
        df = df.reindex(columns=usecols, fill_value="")
        df = df.astype(str).replace(['nan', 'NaN', 'None', '<NA>', 'None'], '')
        
        return df
        
    except Exception as e:
        raise RuntimeError(f"데이터 읽기 실패: {e}")


def write_to_open_excel(book_name: str, sheet_name: str, header_row: int, 
                        df_result: pd.DataFrame, take_cols: List[str], key_cols: List[str],
                        use_color: bool = True):
    """
    매칭 결과를 다시 엑셀에 기입 (xlwings)
    """
    wb = _get_book_by_name(book_name)
    sh = wb.sheets[sheet_name]
    
    # 1. 헤더에서 매칭할 컬럼들(key + take)의 인덱스(1-based) 찾기
    rng_header = sh.range((header_row, 1)).expand('right')
    headers = rng_header.value
    if not isinstance(headers, list): headers = [headers]
    headers = [str(h).strip() for h in headers]
    
    # 컬럼명 -> 인덱스 매핑
    col_map = {h: i+1 for i, h in enumerate(headers)}
    
    # 2. DataFrame 순회하며 값 쓰기?? 
    # -> 너무 느림. 
    # --> 우리는 "원본 순서 유지" & "_idx"가 있다고 가정하거나...
    # 하지만 match_universal 결과인 df_result는 
    #   joined = matched + unmatched (순서가 섞여있을 수도 있음)
    #   하지만 matcher.py에서 `joined = joined.loc[df_b.sort_values("_idx").index]` 등으로 원본 순서 복원 로직이 있나?
    #   matcher.py 코드를 보면:
    #      if use_fast: ... joined = joined.loc[df_b.sort_values("_idx").index]
    #      else: joined = pd.merge(..., how="left") ...
    #   기본적으로 원본 base 순서(행 번호)와 1:1 대응된다고 가정해야 덮어쓰기가 가능.
    #   만약 행이 추가/삭제되었다면 이 방식은 위험함. 
    #   하지만 "매칭" 프로그램 특성상 원본 옆에 붙여넣기를 기대함.
    
    # 입력할 데이터 준비
    # df_result 순서대로 각 행에 기입.
    # 단, 엑셀에 필터가 걸려있거나 숨김 행이 있으면 xlwings range assigment가 위험할 수 있음.
    # 그래도 일반적으로 range().value = ... 가 빠름.
    
    # 쓸 데이터만 뽑기
    vals = df_result[take_cols].values.tolist()
    
    # 쓰기 시작 위치
    start_row = header_row + 1
    
    # 컬럼별로 쓰기 (한 번에 통으로 쓰려면 컬럼들이 연속되어야 하는데, 떨어져 있을 수 있음)
    for col_name in take_cols:
        col_idx = col_map.get(col_name)
        if not col_idx:
            continue # 헤더에 없는 컬럼은 패스 (혹은 맨 뒤에 추가? 지금은 패스)
            
        # 해당 컬럼의 데이터만 추출
        col_data = df_result[col_name].values.tolist()
        # 세로로 쓰기 위해 list of list 변환 [[v], [v], ...]
        col_data_vert = [[v] for v in col_data]
        
        start_cell = sh.range(start_row, col_idx)
        # 범위 지정
        target_rng = sh.range(start_cell, (start_row + len(col_data) - 1, col_idx))
        target_rng.value = col_data_vert
        
        # 색상 강조
        if use_color:
            # 값이 있는 셀만 칠하기? or 전체 칠하기?
            # xlwings color는 (r,g,b). 
            # 한 번에 칠하면 빠름.
            target_rng.color = (255, 255, 204) # 연한 노랑
