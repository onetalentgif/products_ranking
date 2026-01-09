import os
import xlwings as xw
from datetime import datetime, timedelta
from typing import List, Dict, Set, Optional, Tuple, Any
from config import EXCEL_PATH


def get_keyword_from_xlsm() -> Set[str]:
    """
    Extract unique keywords from the 'J' column (10th column) starting from row 7.
    Returns a set of keywords.
    """
    if not os.path.exists(EXCEL_PATH):
        print(f"파일을 찾을 수 없습니다: {EXCEL_PATH}")
        return set()

    app = xw.App(visible=False)
    app.display_alerts = False

    try:
        try:
            wb = app.books.open(EXCEL_PATH, update_links=False)
            
            # Update external links manually
            try:
                links = wb.api.LinkSources(1)
                if links is not None:
                    wb.api.UpdateLink(Name=links)
                    print("외부 연결 데이터 업데이트 완료.")
            except Exception as update_err:
                print(f"외부 링크 업데이트 건너뜀 (사유: {update_err})")

            ws = wb.sheets['데이터']
            last_row = ws.range('J' + str(ws.api.Rows.Count)).end('up').row

            if last_row < 7:
                return set()

            # Batch read J column
            values = ws.range((7, 10), (last_row, 10)).value
            if not isinstance(values, list):
                values = [values]

            keywords = {str(v).strip() for v in values if v}
            print(f"키워드 추출: {list(keywords)}")
            return keywords

        finally:
            if 'wb' in locals():
                wb.close()

    finally:
        app.quit()


def sync_date_columns_until_today(ws, start_date_str: str = "2026-01-01") -> None:
    """
    Ensure date columns exist from start_date up to today.
    Inserts new columns dynamically to the left of fixed fields ('비고', etc.) if missing.
    """
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    today = datetime.now()
    start_search_col = 74  # Column BV

    # Pre-fetch header row (Row 5) once for performance
    max_col = ws.range(5, ws.api.Columns.Count).end('left').column
    if max_col < start_search_col: 
        max_col = start_search_col + 100
        
    row_5_values = ws.range((5, 1), (5, max_col + 10)).value
    
    existing_headers = set()
    for val in row_5_values:
        if val is None:
            continue
        if isinstance(val, datetime):
            existing_headers.add(f"{val.month}/{val.day}")
        else:
            existing_headers.add(str(val).strip())

    current_date = start_date
    while current_date <= today:
        date_header = f"{current_date.month}/{current_date.day}"
        
        if date_header in existing_headers:
             current_date += timedelta(days=1)
             continue

        target_col = start_search_col
        while True:
            val = ws.range(5, target_col).value
            val_str = str(val).strip() if val else ""
            
            if val is None or any(kw in val_str for kw in ["직전", "비고", "서식", "공란"]):
                break
            target_col += 1

        ws.range((1, target_col)).api.EntireColumn.Insert()
        ws.range(5, target_col).value = date_header
        print(f"새 열 추가됨: {target_col}번째 열 ({date_header})")
        
        existing_headers.add(date_header)
        current_date += timedelta(days=1)


def build_row_index(ws) -> Dict[Tuple[str, str], int]:
    """
    Build a map of {(product_id, keyword): row_number}.
    Used for O(1) row lookup during updates.
    """
    last_row = ws.range('J' + str(ws.api.Rows.Count)).end('up').row
    data = ws.range((7, 6), (last_row, 10)).value
    
    row_index = {}
    if not data:
        return row_index

    for i, row in enumerate(data):
        raw_id = row[0]
        kw = row[4]
        
        try:
            vi_id = str(int(float(raw_id))) if raw_id else ""
        except:
            vi_id = str(raw_id or "").strip()

        row_index[(vi_id, str(kw).strip())] = i + 7
        
    return row_index


def map_date_columns(ws) -> Dict[str, int]:
    """
    Map 'YYYY-MM-DD' strings to their Column Index.
    """
    date_to_col = {}
    start_col = 74
    max_col = ws.range(5, ws.api.Columns.Count).end('left').column
    if max_col < start_col:
        return date_to_col

    headers = ws.range((5, start_col), (5, max_col)).value
    if not isinstance(headers, list):
        headers = [headers]

    for i, val in enumerate(headers):
        if val:
            if isinstance(val, datetime):
                d_str = val.strftime('%Y-%m-%d')
            else:
                s_val = str(val).strip()
                if "/" in s_val and len(s_val) <= 5:
                     try:
                        md = datetime.strptime(s_val, "%m/%d")
                        d_str = f"2026-{md.month:02d}-{md.day:02d}"
                     except:
                        d_str = s_val
                else:
                    d_str = s_val

            date_to_col[d_str] = start_col + i

    return date_to_col


def update_product_ranks(ws, row_num: int, date_col_map: Dict[str, int], results: Dict[str, str]) -> None:
    """
    Batch update ranks for a specific product row.
    results: { '2026-01-01': '3', '2026-01-02': '5' }
    """
    updates = []
    for date_str, rank in results.items():
        col = date_col_map.get(date_str)
        if col:
             updates.append((col, rank))

    if not updates:
        return

    updates.sort(key=lambda x: x[0])
    
    is_contiguous = True
    if len(updates) > 1:
        for i in range(len(updates) - 1):
             if updates[i+1][0] != updates[i][0] + 1:
                 is_contiguous = False
                 break
    
    if is_contiguous:
        start_col = updates[0][0]
        values = [u[1] for u in updates]
        ws.range(row_num, start_col).value = values
        print(f" -> {row_num}행 업데이트 완료 (Values: {values})")
    else:
        for col, val in updates:
            ws.range(row_num, col).value = val
            print(f" -> {row_num}행 {col}열 업데이트 ({val})")


def get_dates_requiring_update(ws) -> List[str]:
    """
    Return list of 'YYYY-MM-DD' dates that have NO ranking data in '순위' rows.
    """
    COL_BV = 74
    dates_to_update = []

    max_col = ws.range(5, ws.api.Columns.Count).end('left').column
    if max_col < COL_BV:
        return dates_to_update

    headers = ws.range((5, COL_BV), (5, max_col + 1)).value
    last_row = ws.range('J' + str(ws.api.Rows.Count)).end('up').row

    row_types = ws.range((7, 17), (last_row, 17)).value
    if not isinstance(row_types, list):
         row_types = [row_types]

    for i, header_val in enumerate(headers):
        if header_val is None:
            break
        
        header_str = str(header_val).strip()
        if any(kw in header_str for kw in ["직전", "비고", "서식", "공란"]):
            break

        col_idx = COL_BV + i
        
        formatted_date = ""
        if isinstance(header_val, datetime):
            formatted_date = header_val.strftime('%Y-%m-%d')
        else:
             if "/" in header_str:
                 try:
                     md = datetime.strptime(header_str, "%m/%d")
                     formatted_date = f"2026-{md.month:02d}-{md.day:02d}"
                 except:
                     formatted_date = header_str
             else:
                 formatted_date = header_str

        col_data = ws.range((7, col_idx), (last_row, col_idx)).value
        if not isinstance(col_data, list):
            col_data = [col_data]

        has_actual_rank_data = False
        for idx, rank_val in enumerate(col_data):
            r_type = str(row_types[idx]).strip() if idx < len(row_types) and row_types[idx] else ""
            
            if "순위" in r_type:
                 clean_val = str(rank_val).strip() if rank_val is not None else ""
                 if clean_val not in ["", "0", "0.0", "-", "None"]:
                     has_actual_rank_data = True
                     break
        
        if not has_actual_rank_data:
            dates_to_update.append(formatted_date)

    return dates_to_update