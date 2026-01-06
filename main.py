import logging
import sys
import os
import random
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from openpyxl import load_workbook
from datetime import datetime, timedelta




def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_ROOT_DIR = os.path.join(BASE_DIR, "Chrome_profile")
EXCEL_PATH = os.path.join(BASE_DIR, 'TOP★점프_트래픽관리_순위입력_전.xlsm')
TOP_ADS_URL = "https://top.re.kr/ads"


def create_driver(user_id, headless: bool = False) -> webdriver.Chrome:
    options = Options()
    user_data_path = os.path.join(PROFILE_ROOT_DIR, user_id)
    options.add_argument(f"--user-data-dir={user_data_path}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(1300, 900)
    return driver


def type_like_human(element, text):
    for char in text:
        element.send_keys(char)
        delay = random.uniform(0.2, 0.6)
        time.sleep(delay)


def login_top_with_send_keys(driver, account, debug: bool = True):
    try:
        if debug:
            print(f"로그인 시도: {account}")

        driver.get(TOP_ADS_URL)
        time.sleep(2)

        id_ele = driver.find_element(By.XPATH, '//label[contains(text(), "아이디")]/following-sibling::input')
        pw_ele = driver.find_element(By.XPATH, '//label[contains(text(), "비밀번호")]/following-sibling::input')

        # ID 한 글자씩 타이핑
        id_ele.clear()
        time.sleep(0.5)

        id_ele.click()
        type_like_human(id_ele, account["user_id"])
        time.sleep(1)

        # PW 한 글자씩 타이핑
        pw_ele.click()
        type_like_human(pw_ele, account["user_pw"])
        time.sleep(1)

        # 로그인 상태 유지 버튼 클릭
        keep_ele = driver.find_element(By.ID, "remember")
        keep_ele.click()
        time.sleep(1)

        # 로그인 버튼 클릭
        login_btn = driver.find_element(By.XPATH, '//button[contains(text(), "로그인")]')
        login_btn.click()
        time.sleep(3)

        return True

    except Exception as e:
        if debug:
            print(f"[LOGIN] 로그인 중 예외: {e}")
        return False


def is_top_logged_in(driver, timeout: int = 3) -> bool:

    driver.get("https://top.re.kr/ads")
    wait = WebDriverWait(driver, timeout)

    # 1) 로그아웃 버튼이 '보이면' 로그인 상태로 판단
    try:
        logout_btn = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//button[contains(text(), "로그아웃")]')
            )
        )
        if logout_btn.is_displayed():
            print("로그아웃 버튼이 보여서, 로그인 상태로 판단합니다.")
            return True
    except TimeoutException:
        pass
    except Exception as e:
        print("로그아웃 버튼 확인 중 예외 발생:", e)

    # 2) 로그인 버튼이 '보이면' 로그아웃 상태로 판단
    try:
        login_btn = driver.find_element(By.XPATH, '//button[contains(text(), "로그인")]')
        if login_btn.is_displayed():
            print("로그인 버튼이 보여서, 아직 로그인 안 된 상태로 판단합니다.")
            return False
    except Exception as e:
        print("로그인 버튼 확인 중 예외 발생:", e)

    # 3) 둘 다 못 찾았으면 애매한 상태 → 보수적으로 '로그아웃 상태'로 가정
    print("로그인/로그아웃 버튼을 찾지 못했습니다. 로그아웃 상태로 간주합니다.")
    return False


def login_success_check(driver, account):
    user_id = account['user_id']

    try:
        if is_top_logged_in(driver, 3):
            print(f"[{user_id}] 이미 로그인되어 있습니다.")
            return True
    except Exception as e:
        print(f"로그인 상태 확인 중 에러: {e}")

    print(f"[{user_id}] 로그인 세션이 없습니다. 로그인을 시도합니다.")

    # 한 계정 당 최대 로그인 재시도 횟수 설정
    MAX_RETRIES = 3

    for attempt in range(1, MAX_RETRIES + 1):
        print(f"로그인 시도 {attempt}/{MAX_RETRIES}회 진행 중...")

        try:
            login_top_with_send_keys(driver, account)

            # 로그인 성공 여부 확인
            if is_top_logged_in(driver, 5):
                print(f"[{user_id}] 로그인 성공! 작업을 시작합니다.")
                return True

        except Exception as e:
            print(f"[{user_id}] 시도 중 에러 발생: {e}")

        if attempt < MAX_RETRIES:
            print(f"[{user_id}] 로그인 실패. {attempt+1}회차 재시도를 위해 대기합니다.")
            time.sleep(3)  # 너무 빠른 재시도는 차단 위험이 있음
        else:
            print(f"[{user_id}] 모든 재시도 횟수를 소진했습니다.")

    print(f"[{user_id}] 최종적으로 로그인에 실패하였습니다.")
    return False


def search_keyword(driver, keyword: str, timeout: int = 10):
    try:
        wait = WebDriverWait(driver, timeout)

        search_input = wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, '슬롯번호, 아이디, 키워드')]")))

        search_input.clear()
        search_input.send_keys(keyword)

        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '검색')]")))
        driver.execute_script("arguments[0].click();", search_button)

        time.sleep(2)
        print(f"'{keyword}' 검색 완료")

    except Exception as e:
        print(f"키워드 검색 중 오류 발생: {e}")


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


def extract_product_results(driver, target_dates: list, timeout: int = 10):
    wait = WebDriverWait(driver, timeout)

    product_results = {d_text: [] for d_text in target_dates}

    # 비교를 위해 타겟 날짜들을 datetime 객체로 변환
    target_dt_list = [datetime.strptime(d, '%Y-%m-%d') for d in target_dates]

    try:
        # [수정] 테이블의 데이터(tr)가 최소한 하나라도 나타날 때까지 기다림
        # 검색 후 로딩 시간을 고려하여 확실하게 대기합니다.
        time.sleep(1.5)

        try:
            # tbody 안에 tr이 있는지 확인
            rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))
        except TimeoutException:
            print("테이블 응답 시간 초과: 결과가 없거나 로딩이 너무 느립니다.")
            return product_results

        # '조회 결과 없음' 문구가 있는 경우 처리
        if len(rows) == 0 or (len(rows) == 1 and "정보가 없습니다" in rows[0].text):
            print("조회 결과 없음 (표시된 데이터가 없습니다)")
            return product_results

        for row in rows:
            try:
                # 시작일(12)과 종료일(13) 추출
                start_date_text = row.find_element(By.XPATH, "./td[12]").text.strip()
                end_date_text = row.find_element(By.XPATH, "./td[13]").text.strip()

                if not start_date_text or not end_date_text:
                    continue

                start_date = datetime.strptime(start_date_text, '%Y-%m-%d')
                end_date = datetime.strptime(end_date_text, '%Y-%m-%d')

                # 타겟 날짜 중 가장 오래된 날짜보다 현재 데이터의 종료일이 더 과거라면 건너뜀
                if end_date < min(target_dt_list):
                    continue

                # 순위(9) 및 ID(8) 추출
                rank_text = row.find_element(By.XPATH, "./td[9]").text.strip()
                rank_number = ""
                if "순위밖" not in rank_text and "위" in rank_text:
                    rank_number = rank_text.split('위')[0].strip()

                url = row.find_element(By.XPATH, "./td[8]//a").get_attribute("href")
                last_id = url.split("=")[-1]

                # 모든 타겟 날짜에 대해 매칭 확인
                for i, target_dt in enumerate(target_dt_list):
                    if start_date <= target_dt <= end_date:
                        product_results[target_dates[i]].append((last_id, rank_number))
                        print(f"매칭 발견: {target_dates[i]} | ID: {last_id} | 순위: {rank_number}")

            except Exception:
                # 개별 행 파싱 실패 시 다음 행으로 진행
                continue

    except Exception as e:
        print(f"테이블 처리 중 오류 발생: {e}")

    return product_results


# 2026-01-01부터 날짜 열 추가
def sync_date_columns_until_today(ws, start_date_str="2026-01-01"):
    """
    1월 1일부터 오늘까지 누락된 날짜 열을 5행에 자동으로 추가
    고정 필드를 유지하기 위해 열을 삽입(Insert)하며 확장
    """
    # 날짜 범위 설정
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    today = datetime.now()

    # 고정 열(BW열 등) 시작 위치
    fixed_start_col = 75

    # 1월 1일부터 오늘까지 반복하며 날짜 확인
    current_date = start_date
    while current_date <= today:
        date_header = f"{current_date.month}/{current_date.day}"

        # 5행에서 해당 날짜가 이미 있는지 확인
        is_exist = False

        cell_val = ws.cell(row=5, column=74).value

        # 날짜 객체 또는 문자열 비교
        if isinstance(cell_val, datetime):
            check_str = f"{cell_val.month}/{cell_val.day}"
        else:
            check_str = str(cell_val).strip()

        if check_str == date_header:
            is_exist = True

        # 날짜가 없으면 열을 삽입하고 날짜 입력
        if not is_exist:
            # 고정 열 바로 앞에 새 열 삽입
            ws.insert_cols(fixed_start_col)
            # 삽입된 위치에 날짜 기입
            ws.cell(row=5, column=fixed_start_col).value = date_header
            print(f"새 열 추가됨: {fixed_start_col}번째 열 ({date_header})")

            # 열이 하나 추가되었으므로 고정 열의 위치(기준점)도 오른쪽으로 한 칸 이동
            fixed_start_col += 1

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
        if "직전" in cell_str or "비고" in cell_str or "서식" in cell_str:
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


if __name__ == "__main__":
    account = {"user_id": "sstrade251016", "user_pw": "a2345"}
    main_wb = load_workbook(EXCEL_PATH, keep_vba=True)
    main_ws = main_wb['데이터']

    # 1. 날짜 동기화
    sync_date_columns_until_today(main_ws, start_date_str="2026-01-01")
    keywords = get_keyword_from_xlsm()

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

    driver = create_driver(account["user_id"], headless=False)

    try:
        if login_success_check(driver, account):
            for keyword in keywords:
                # 'end' 키워드 등이 섞여 있을 경우를 대비한 방어 로직
                if "end" in keyword.lower(): continue

                missing_dates = get_missing_dates_for_keyword(main_ws, keyword, date_info_map)
                if not missing_dates:
                    print(f"[{keyword}] 입력할 칸이 없습니다. 패스.")
                    continue

                search_keyword(driver, keyword)
                all_results_map = extract_product_results(driver, missing_dates)

                # 결과 데이터가 있을 때만 엑셀 기록 루프 실행
                if all_results_map:
                    for date_text, product_results in all_results_map.items():
                        for p_id, p_rank in product_results:
                            update_excel_rank(main_ws, p_id, keyword, p_rank, date_text)

            main_wb.save(EXCEL_PATH)
            print("\n파일 저장 완료!")

    except Exception as e:
        print(f"실행 중 오류 발생: {e}")
    finally:
        main_wb.close()
        driver.quit()