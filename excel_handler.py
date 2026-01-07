import os
from openpyxl import load_workbook
from datetime import datetime, timedelta
from config import EXCEL_PATH


def get_keyword_from_xlsm():
    if not os.path.exists(EXCEL_PATH):
        print(f"파일을 찾을 수 없습니다: {EXCEL_PATH}")
        return set()

    # 1. 엑셀 로드
    # keep_vba=True: 매크로 유지
    # data_only=True: 수식이 아닌 '텍스트 결과값'만 가져옴
    temp_wb = load_workbook(EXCEL_PATH, keep_vba=True, data_only=True)
    ws = temp_wb['데이터']

    # 2. 키워드 가져오기
    keywords = set()
    for row in ws.iter_rows(min_row=7, min_col=10, max_col=10):
        cell_value = row[0].value # iter_rows는 한 행을 셀들의 묶음(튜플)으로 반환

        if cell_value is None:
            continue

        keyword = str(cell_value).strip()

        if not keyword:
            continue

        keywords.add(keyword)

    print(f"키워드 추출: {list(keywords)}")

    temp_wb.close()
    return keywords



# 2026-01-01부터 날짜 열 확인 및 추가
def sync_date_columns_until_today(ws, start_date_str="2026-01-01"):
    """
    1월 1일부터 오늘까지 누락된 날짜 열을 5행에 자동으로 추가
    고정 필드를 유지하기 위해 열을 삽입(Insert)하며 확장
    """
    # 날짜 범위 설정
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    today = datetime.now()

    # 날짜를 넣기 시작할 시작 열 (BV = 74)
    start_search_col = 74

    # 1월 1일부터 오늘까지 반복하며 날짜 확인
    current_date = start_date
    while current_date <= today:
        date_header = f"{current_date.month}/{current_date.day}"
        is_exist = False    # 5행에서 해당 날짜가 이미 있는지 확인

        # cell_val = ws.cell(row=5, column=74).value

        # 현재 5행에 해당 날짜가 이미 있는지 끝까지 검사
        last_col = ws.max_column
        for col in range(start_search_col, last_col+1):
            cell_val = ws.cell(row=5, column=col).value
            if cell_val is None:
                continue

            # 날짜 (객체 또는 문자열) 비교
            if isinstance(cell_val, datetime):
                check_str = f"{cell_val.month}/{cell_val.day}"
            else:
                check_str = str(cell_val).strip()

            if check_str == date_header:
                is_exist = True
                break

        # 날짜가 없으면 "가장 오른쪽 끝" 혹은 "특정 고정필드 직전"에 추가
        if not is_exist:
            # '비고'나 '서식' 같은 고정 헤더가 시작되는 위치를 찾음
            target_col = start_search_col
            while True:
                val = ws.cell(row=5, column=target_col).value
                # 빈칸이거나 고정 키워드가 나오면 그 자리에 삽입
                if val is None or any(kw in str(val) for kw in ["직전", "비고", "서식", "공란"]):
                    break
                target_col += 1

            ws.insert_cols(target_col)
            ws.cell(row=5, column=target_col).value = date_header
            print(f"새 열 추가됨: {target_col}번째 열 ({date_header})")

        current_date += timedelta(days=1)



def get_all_date_texts_from_header(ws):
    """
    엑셀 5행의 BV열(74)부터 고정 필드 전까지의 날짜들을
    ['2026-01-01', '2026-01-02', ...] 형태의 텍스트 리스트로 반환
    """
    COL_BV = 74
    date_text_list = []

    # BV열(74)부터 오른쪽으로 하나씩 검사
    for col in range(COL_BV, ws.max_column + 1):
        cell_val = ws.cell(row=5, column=col).value

        # 빈 칸이거나 날짜가 아닌 고정 헤더를 만나면 탐색 중단
        if cell_val is None:
            break

        cell_str = str(cell_val).strip()
        if any(kw in cell_str for kw in ["직전", "비고", "서식", "공란"]):
            break

        # 날짜 데이터를 텍스트로 변환
        formatted_date = None

        # 셀 값이 datetime 객체인 경우
        if isinstance(cell_val, datetime):
            formatted_date = cell_val.strftime('%Y-%m-%d')

        # 셀 값이 '1/1' 또는 '2026-01-01' 같은 문자열인 경우
        else:
            try:
                # 이미 'YYYY-MM-DD' 형식인지 확인
                datetime.strptime(cell_str, '%Y-%m-%d')
                formatted_date = cell_str
            except ValueError:
                # '1/1' 형식인 경우 현재 연도를 붙여서 변환
                try:
                    current_year = datetime.now().year
                    dt = datetime.strptime(f"{current_year}/{cell_str}", "%Y/%m/%d")
                    formatted_date = dt.strftime('%Y-%m-%d')
                except:
                    # 형식을 알 수 없으면 그대로 문자열로 저장
                    formatted_date = cell_str

        if formatted_date:
            date_text_list.append(formatted_date)

    return date_text_list


def get_missing_dates_for_keyword(ws, keyword, date_info_map):
    """지정된 키워드에 대해 순위값이 비어있는 날짜들만 반환"""
    missing_dates = set()
    for row in range(7, ws.max_row + 1):
        ex_kw = str(ws.cell(row=row, column=10).value or "").strip()
        if ex_kw == keyword:
            for d_str, col_idx in date_info_map.items():
                if ws.cell(row=row, column=col_idx).value is None:
                    missing_dates.add(d_str)
    return list(missing_dates)


def update_excel_rank(ws, target_vi_id, target_keyword, rank_value, target_date):
    # 5행에서 날짜에 해당하는 열 번호 찾기
    target_col = None
    # 날짜 입력 형식을 안전하게 변환
    dt = datetime.strptime(target_date, '%Y-%m-%d')
    search_header = f"{dt.month}/{dt.day}"

    for col in range(74, ws.max_column + 1):
        cell_val = ws.cell(row=5, column=col).value
        if cell_val is None: continue

        # 엑셀 헤더가 날짜 객체인지 텍스트인지 판별하여 비교
        header = f"{cell_val.month}/{cell_val.day}" if isinstance(cell_val, datetime) else str(cell_val).strip()

        if header == search_header:
            target_col = col
            break

    if not target_col:
        print(f"엑셀에서 {search_header} 열을 찾지 못했습니다.")
        return

    COL_VI_ID = 6  # F열
    COL_KEYWORD = 10  # J열

    found = False
    for row in range(7, ws.max_row + 1):
        # [수정] ID 비교 시 float 형태(.0)가 생기지 않도록 정수형 처리 후 문자열 변환
        raw_vi_id = ws.cell(row=row, column=COL_VI_ID).value
        try:
            # 12345.0 같은 데이터를 '12345'로 변환
            vi_id = str(int(float(raw_vi_id))) if raw_vi_id else ""
        except:
            vi_id = str(raw_vi_id or "").strip()

        keyword = str(ws.cell(row=row, column=COL_KEYWORD).value or "").strip()

        # ID와 키워드 동시 비교
        if vi_id == str(target_vi_id) and keyword == target_keyword:
            ws.cell(row=row, column=target_col).value = rank_value
            print(f"성공: {row}행 {target_col}열에 '{rank_value}' 입력")
            found = True
            break

    if not found:
        # 디버깅을 위해 찾지 못한 정보 출력
        pass


def get_dates_requiring_update(ws):
    """
    5행의 날짜 열들을 순회하며, 해당 열 전체(7행~마지막행)에
    순위 데이터가 '단 하나도' 없는 날짜 리스트만 반환합니다.
    """
    COL_BV = 74
    dates_to_update = []

    for col in range(COL_BV, ws.max_column + 1):
        header_val = ws.cell(row=5, column=col).value
        if header_val is None: break

        header_str = str(header_val).strip()
        # 고정 헤더를 만나면 탐색 중단
        if any(kw in header_str for kw in ["직전", "비고", "서식", "공란"]):
            break

        # --- 수정된 로직 시작 ---
        has_any_data = False
        for row in range(7, ws.max_row + 1):
            rank_val = ws.cell(row=row, column=col).value

            # 셀 값이 None이 아니고, 공백 제외 문자열이 존재하면 데이터가 있는 것으로 간주
            if rank_val is not None and str(rank_val).strip() != "":
                has_any_data = True
                break  # 하나라도 데이터가 있으면 이 날짜는 검사 종료

        # 데이터가 '하나도 없는' 날짜만 업데이트 대상으로 선정
        if not has_any_data:
            # 날짜 형식 통일 (YYYY-MM-DD)
            if isinstance(header_val, datetime):
                formatted_date = header_val.strftime('%Y-%m-%d')
            else:
                try:
                    current_year = datetime.now().year
                    dt = datetime.strptime(f"{current_year}/{header_str}", "%Y/%m/%d")
                    formatted_date = dt.strftime('%Y-%m-%d')
                except:
                    formatted_date = header_str

            dates_to_update.append(formatted_date)
        # --- 수정된 로직 끝 ---

    if dates_to_update:
        print(f"데이터가 완전히 비어 있는 날짜(크롤링 대상): {dates_to_update}")
    else:
        print("모든 날짜 열에 최소 하나 이상의 데이터가 기록되어 있습니다.")

    return dates_to_update