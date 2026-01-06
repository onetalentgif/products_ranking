from config import EXCEL_PATH, ACCOUNT
import web_handler as wh
import excel_handler as eh
from openpyxl import load_workbook
from datetime import datetime


if __name__ == "__main__":
    account = {"user_id": "sstrade251016", "user_pw": "a2345"}
    main_wb = load_workbook(EXCEL_PATH, keep_vba=True)
    main_ws = main_wb['데이터']

    # 1. 날짜 동기화
    eh.sync_date_columns_until_today(main_ws, start_date_str="2026-01-01")
    keywords = eh.get_keyword_from_xlsm()

    # 2. 날짜-열 매핑 생성 (YYYY-MM-DD 자릿수 엄격 적용)
    date_info_map = {}
    for col in range(74, main_ws.max_column + 1):
        val = main_ws.cell(row=5, column=col).value
        if val is None: break

        val_str = str(val).strip()
        try:
            # 어떠한 형식이든 datetime 객체로 바꾼 뒤 '2026-01-06' 형태로 통일
            if isinstance(val, datetime):
                d_str = val.strftime('%Y-%m-%d')
            else:
                d_str = datetime.strptime(f"2026/{val_str}", "%Y/%m/%d").strftime('%Y-%m-%d')
            date_info_map[d_str] = col
        except ValueError:
            print(f"날짜가 아닌 헤더 발견 후 중단: {val_str}")
            break

    driver = wh.create_driver(account["user_id"], headless=False)

    try:
        if wh.ogin_success_check(driver, account):
            for keyword in keywords:
                # 'end' 키워드 등이 섞여 있을 경우를 대비한 방어 로직
                if "end" in keyword.lower(): continue

                missing_dates = eh.get_missing_dates_for_keyword(main_ws, keyword, date_info_map)
                if not missing_dates:
                    print(f"[{keyword}] 입력할 칸이 없습니다. 패스.")
                    continue

                wh.search_keyword(driver, keyword)
                all_results_map = wh.extract_product_results(driver, missing_dates)

                # 결과 데이터가 있을 때만 엑셀 기록 루프 실행
                if all_results_map:
                    for date_text, product_results in all_results_map.items():
                        for p_id, p_rank in product_results:
                            eh.update_excel_rank(main_ws, p_id, keyword, p_rank, date_text)

            main_wb.save(EXCEL_PATH)
            print("\n파일 저장 완료!")

    except Exception as e:
        print(f"실행 중 오류 발생: {e}")
    finally:
        main_wb.close()
        driver.quit()