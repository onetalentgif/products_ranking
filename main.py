import os
import sys
import traceback
from datetime import datetime
import xlwings as xw
from typing import List, Dict, Set, Optional

from config import EXCEL_PATH, ACCOUNT, MAX_PAGES
from excel_handler import (
    sync_date_columns_until_today,
    get_dates_requiring_update,
    build_row_index,
    map_date_columns,
    update_product_ranks
)
from web_handler import (
    create_driver,
    login_success_check,
    search_keyword,
    extract_product_results,
    delete_chrome_cache,
    kill_chrome_processes,
    extract_normal_products
)


class RankingAutomation:
    def __init__(self):
        self.app: Optional[xw.App] = None
        self.wb: Optional[xw.Book] = None
        self.ws: Optional[xw.Sheet] = None
        self.driver = None

    def initialize_excel(self):
        """Initialize Excel application and open the workbook."""
        if not os.path.exists(EXCEL_PATH):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {EXCEL_PATH}")

        print("엑셀 파일을 불러오는 중입니다...")
        
        # [수정] exe 실행 여부 감지
        is_frozen = getattr(sys, "frozen", False)
        
        # exe 실행이면 보이지 않게(False), 개발 모드면 보이게(True)
        self.app = xw.App(visible=not is_frozen)
        self.app.display_alerts = False
        
        # Open without updating links initially to avoid popups
        self.wb = self.app.books.open(EXCEL_PATH, update_links=False)
        
        # Manually update links
        try:
            links = self.wb.api.LinkSources(1)
            if links:
                self.wb.api.UpdateLink(Name=links)
                print("외부 연결 데이터 업데이트 완료.")
        except Exception as e:
            print(f"외부 링크 업데이트 건너뜀: {e}")

        self.ws = self.wb.sheets['데이터']

    def initialize_browser(self):
        """Prepare environment and start Chrome."""
        print("브라우저 환경 준비 중...")
        # delete_chrome_cache(ACCOUNT["user_id"])  # Optional: clear cache if needed
        
        # [수정] exe 실행 여부 감지
        is_frozen = getattr(sys, "frozen", False)

        # exe 실행이면 headless(화면 없음) 모드 활성화
        self.driver = create_driver(ACCOUNT["user_id"], headless=is_frozen)

    def run(self):
        try:
            # 1. Excel Prep
            self.initialize_excel()
            
            # Sync dates and check work
            sync_date_columns_until_today(self.ws)
            target_dates = get_dates_requiring_update(self.ws)
            
            if not target_dates:
                print(">>> 모든 날짜에 데이터가 이미 존재합니다. 추가로 작업할 내용이 없습니다.")
                return

            print(f">>> 수집 대상 날짜: {target_dates}")

            # Pre-calculate indices
            row_index = build_row_index(self.ws)
            date_col_map = map_date_columns(self.ws)

            # 2. Browser Prep & Login
            self.initialize_browser()
            if not login_success_check(self.driver, ACCOUNT):
                print("로그인 실패로 작업을 중단합니다.")
                return

            # 3. Discovery Phase
            print("\n>>> '정상' 상품 리스트에서 키워드 추출 중...")
            normal_data = extract_normal_products(self.driver, target_dates)
            
            if not normal_data:
                print("수집된 정상 상품 데이터가 없습니다.")
                return

            # Extract unique Product IDs
            product_ids = set()
            for items in normal_data.values():
                for item in items:
                    # item structure: (keyword, product_id, rank)
                    if item[1]:
                        product_ids.add(item[1])

            print(f"-> 추출된 고유 상품 번호: {len(product_ids)}개 ({list(product_ids)})")

            # 4. Main Scraping Loop
            print("\n>>> 상품 번호별 상세 검색 및 엑셀 업데이트 시작")
            
            for p_id in product_ids:
                print(f"\n--- 현재 검색 상품 번호: [{p_id}] ---")
                
                # Search by Product ID
                if not search_keyword(self.driver, p_id):
                    continue

                # Extract results
                product_results = extract_product_results(self.driver, target_dates)
                # Structure: { datetime_obj: [(kw, pid, rank), ...], ... }

                # 5. Process & Batch Update per Product ID
                # We need to map scraping results back to the Excel rows
                # One Product ID search might return multiple rows (different keywords) for different dates.
                
                # We need to aggregate results by (Product ID, Keyword) -> { Date: Rank }
                # Because our Excel Row Index is based on (ID, Keyword).
                
                aggregated_updates = {} # Key: (id, kw), Value: { date_str: rank }

                found_count = 0
                for date_obj, items in product_results.items():
                    date_str = date_obj.strftime('%Y-%m-%d')
                    
                    if not items:
                        # print(f" [{date_str}] 결과 없음")
                        continue
                        
                    for item in items:
                        kw = item[0]
                        pid = item[1]
                        rank = item[2]
                        
                        # Only update if it matches the p_id we are searching (sanity check)
                        if pid == p_id:
                            key = (pid, kw)
                            if key not in aggregated_updates:
                                aggregated_updates[key] = {}
                            
                            aggregated_updates[key][date_str] = rank
                            found_count += 1

                # Apply Batch Updates
                if not aggregated_updates:
                    print(f"-> [{p_id}] 유효한 데이터 없음.")
                    continue

                for (pid, kw), date_rank_map in aggregated_updates.items():
                    # Find row number
                    row_num = row_index.get((pid, kw))
                    if not row_num:
                        print(f"엑셀에서 행을 찾을 수 없음: ID={pid}, KW={kw}")
                        continue
                        
                    # Perform Batch Update for this row
                    update_product_ranks(self.ws, row_num, date_col_map, date_rank_map)

                print(f"-> [{p_id}] 처리 완료 (총 {found_count}건 데이터)")

            # 6. Save
            print("\n데이터 기록 완료. 엑셀 파일을 저장합니다...")
            self.wb.save()
            print("저장이 완료되었습니다.")

        except Exception as e:
            print("\n[CRITICAL ERROR] 실행 중 치명적인 오류 발생:")
            traceback.print_exc()
            
        finally:
            self.cleanup()

    def cleanup(self):
        """Close resources safely."""
        print("자원 정리 중...")
        if self.wb:
            try:
                self.wb.close()
                print("엑셀 파일 닫힘.")
            except:
                pass
        
        if self.app:
            try:
                self.app.quit()
                print("엑셀 앱 종료.")
            except:
                pass

        if self.driver:
            try:
                self.driver.quit()
                print("브라우저 종료.")
            except:
                pass


if __name__ == "__main__":
    automation = RankingAutomation()
    automation.run()