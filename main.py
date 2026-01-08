import os
import xlwings as xw
from openpyxl import load_workbook
from config import EXCEL_PATH, ACCOUNT  # ACCOUNT는 {"user_id": "...", "user_pw": "..."} 형태
from excel_handler import (
    get_keyword_from_xlsm,
    sync_date_columns_until_today,
    get_all_date_texts_from_header,
    update_excel_rank, get_dates_requiring_update
)
from web_handler import (
    create_driver,
    login_success_check,
    search_keyword,
    extract_product_results,
    delete_chrome_cache,
    kill_chrome_processes
)


def main():
    # 엑셀 파일 로드
    if not os.path.exists(EXCEL_PATH):
        print(f"파일을 찾을 수 없습니다: {EXCEL_PATH}")
        return

    # [수정] 키워드 목록 가져오기 (excel_handler 내에서 별도로 파일을 열고 닫음)
    keywords = get_keyword_from_xlsm()
    if not keywords:
        print("검색할 키워드가 엑셀에 없습니다.")
        return

    print("엑셀 파일을 불러오는 중입니다...")

    # [수정] xlwings App 인스턴스 생성 및 설정
    app = xw.App(visible=True)  # 작업 과정을 보고 싶다면 True 유지
    app.display_alerts = False  # 엑셀 알림(팝업) 차단

    # wb = load_workbook(EXCEL_PATH, keep_vba=True)
    # ws = wb['데이터']

    # [수정] 팝업 없이 열기
    wb = app.books.open(EXCEL_PATH, update_links=False)

    # [수정] 강제로 외부 링크 최신 데이터로 업데이트
    try:
        links = wb.api.LinkSources(1)
        if links is not None:
            wb.api.UpdateLink(Name=links)
            print("외부 연결 데이터 업데이트 완료.")
        else:
            print("업데이트할 외부 링크가 없습니다.")
    except Exception as update_err:
        print(f"외부 링크 업데이트를 수행할 수 없습니다: {update_err}")

    ws = wb.sheets['데이터']  # [수정] xlwings 시트 선택 방식

    try:
        # 날짜 열 동기화 (오늘 날짜까지 열이 없으면 생성)
        sync_date_columns_until_today(ws)

        # 데이터가 '전혀' 입력되지 않은 날짜 리스트만 추출
        target_dates = get_dates_requiring_update(ws)
        if not target_dates:
            print(">>> 모든 날짜에 데이터가 이미 존재합니다. 추가로 작업할 내용이 없습니다.")
            # wb.close()
            return

        print(f">>> 다음 날짜들에 대해 수집을 시작합니다: {target_dates}")

        # # 키워드 목록 가져오기
        # keywords = get_keyword_from_xlsm()
        # if not keywords:
        #     print("검색할 키워드가 엑셀에 없습니다.")
        #     wb.close()
        #     return
        #
        # # print(f"대상 키워드: {list(keywords)}")

        # # [추가] 작업 시작 전 깨끗하게 정리
        # kill_chrome_processes()
        # 브라우저 실행 전 크롬 캐시 삭제
        delete_chrome_cache(ACCOUNT["user_id"])
        # 브라우저 실행 및 로그인
        driver = create_driver(ACCOUNT["user_id"], headless=False)

        try:
            if login_success_check(driver, ACCOUNT):
                # 각 키워드별 검색 및 데이터 추출
                for keyword in keywords:
                    print(f"\n>>> 키워드 검색 시작: {keyword}")
                    search_keyword(driver, keyword)

                    # 웹 페이지에서 결과 추출 (딕셔너리 형태: {datetime: [(kw, id, rank), ...]})
                    product_results = extract_product_results(driver, target_dates)

                    # 추출된 결과를 엑셀 메모리에 업데이트
                    for target_date, items in product_results.items():
                        date_str = target_date.strftime('%Y-%m-%d')  # 출력용 날짜 변환

                        # product_results가 비었을 때
                        if not items:
                            print(f" [{date_str}] '{keyword}'에 대한 검색 결과가 없습니다.")
                            continue

                        # 각 날짜에 포함된 상품 리스트를 순회 (row_keyword, product_id, rank_number)
                        for item in items:
                            product_keyword = item[0]  # 튜플의 첫 번째: 행 키워드
                            product_id = item[1]  # 튜플의 두 번째: ID
                            product_rank = item[2]  # 튜플의 세 번째: 순위

                            # 엑셀의 해당 날짜/키워드/ID 행을 찾아 순위 입력
                            update_excel_rank(ws, product_id, product_keyword, product_rank, date_str)

            # 작업이 끝난 후 한 번에 저장
            print("\n데이터 기록 완료. 엑셀 파일을 저장합니다...")
            # wb.save(EXCEL_PATH)

            # [수정] 현재 열린 파일에 그대로 저장 (매크로 유지)
            wb.save()
            print("저장이 완료되었습니다.")

        except Exception as e:
            print(f"브라우저 작업 중 오류 발생: {e}")
        finally:
            # wb.close()
            driver.quit()

    except Exception as e:
            print(f"엑셀 처리 중 오류 발생: {e}")
    finally:
        # [수정] 작업 완료 후 반드시 엑셀을 닫아 프로세스 점유 방지
        wb.close()
        print("엑셀 연결을 종료했습니다.")
        # app.quit() # 만약 App(visible=False)를 썼다면 앱 자체도 종료해야 함

if __name__ == "__main__":
    main()