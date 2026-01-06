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


# target_date는 엑셀에 순위 입력한 날짜(예.1/1)를 2026-01-01로 변환한 값
def extract_product_results(driver, target_dates: list, timeout: int = 10):
    wait = WebDriverWait(driver, timeout)

    # 결과를 담을 딕셔너리 (날짜별로 결과 리스트를 저장)
    product_results = {d_text: [] for d_text in target_dates}

    # 비교를 위해 타겟 날짜들을 datetime 객체로 미리 변환
    target_dt_list = [datetime.strptime(d, '%Y-%m-%d') for d in target_dates]

    try:
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))

        if not rows or (len(rows) == 1 and "슬롯 정보가 없습니다" in rows[0].text):
            print("조회 결과 없음")
            return {}

        for row in rows:
            try:
                start_date_text = row.find_element(By.XPATH, "./td[12]").text.strip()
                end_date_text = row.find_element(By.XPATH, "./td[13]").text.strip()

                # 날짜 형식이 빈 값이거나 유효하지 않을 경우 건너뜀
                if not start_date_text or not end_date_text:
                    continue

                # 웹의 날짜 텍스트를 datetime 객체로 변환
                start_date = datetime.strptime(start_date_text, '%Y-%m-%d')
                end_date = datetime.strptime(end_date_text, '%Y-%m-%d')

                rank_text = row.find_element(By.XPATH, "./td[9]").text.strip()
                rank_number = ""
                if "순위밖" in rank_text:
                    rank_number = ""
                elif "위" in rank_text:
                    # "위" 앞에 숫자가 있는 경우만 추출
                    rank_number = rank_text.split('위')[0].strip()

                url = row.find_element(By.XPATH, "./td[8]//a").get_attribute("href")
                last_id = url.split("=")[-1]

                for i, target_date in enumerate(target_dt_list):
                    # 타겟 날짜가 시작~종료 범위 안에 있는 경우
                    if start_date <= target_date <= end_date:
                        is_date_found = True  # 일치하는 날짜 찾음
                        product_results[target_dates[i]].append((last_id, rank_number))
                        print(f"시작일: {start_date}, VI ID: {last_id}, 순위: {rank_number}")

                    # 타겟 날짜가 종료일보다 클때 (이미 지나간 과거 데이터)
                    # 리스트가 최신순 정렬이므로, 이후의 행들은 검사할 필요 없음
                    elif target_date > end_date:
                        print(f"과거 데이터 구간({end_date_text}) 진입. 검사를 중단합니다.")
                        break

                    # 타겟 날짜가 시작일보다 작을때 (아직 시작되지 않은 미래 데이터)
                    # 미래 예약 슬롯은 리스트 상단에 있을 수 있으므로, break 하지 않고 다음 행을 확인
                    elif target_date < start_date:
                        print(f"미래 데이터({start_date_text}) 발견. 건너뛰고 다음 행을 확인합니다.")
                        continue

            except Exception as e:
                continue

        # 날짜는 찾았으나 결과가 비었을 때만 경고
        if not product_results:
            if is_date_found:
                print(f"'{target_date}' 행은 찾았으나, 내부 데이터(VI ID/순위) 추출에 실패했습니다.")

    except Exception as e:
        print(f"테이블 로딩 중 오류 발생: {e}")

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
    # '2026-01-06' -> '1/6' 형식으로 변환하여 비교 준비
    dt = datetime.strptime(target_date, '%Y-%m-%d')
    search_header = f"{dt.month}/{dt.day}"

    # BV열(74번)부터 실제 데이터가 있는 마지막 열까지 탐색
    for col in range(74, ws.max_column + 1):
        cell_val = ws.cell(row=5, column=col).value

        # 날짜 객체 또는 문자열 비교
        if isinstance(cell_val, datetime):
            header = f"{cell_val.month}/{cell_val.day}"
        else:
            header = str(cell_val).strip() if cell_val else ""

        if header == search_header:
            target_col = col
            break

    if not target_col:
        print(f"엑셀에서 날짜 {target_date} ({search_header}) 열을 찾을 수 없습니다.")
        return


    # 열 번호 설정
    COL_VI_ID = 6  # F열
    COL_KEYWORD = 10  # J열

    found = False
    # 7행부터 마지막 행까지 탐색
    for row in range(7, ws.max_row + 1):
        # 엑셀의 VI ID가 숫자형일 수 있으므로 문자열로 변환하여 비교
        vi_id = str(ws.cell(row=row, column=COL_VI_ID).value or "").strip()
        keyword = str(ws.cell(row=row, column=COL_KEYWORD).value or "").strip()

        # 두 조건이 일치하는 행 찾기
        if vi_id == str(target_vi_id) and keyword == target_keyword:
            # 찾은 날짜 열(target_col)에 순위 값 입력
            ws.cell(row=row, column=target_col).value = rank_value
            print(f"{row}행에 {target_col}번 열에 순위 값 '{rank_value}' 입력")
            found = True
            break  # 찾았으므로 루프 종료 (1개 밖에 없는 게 맞는지 확인 필요)

    if not found:
        print(f"{target_date} 조건에 맞는 행을 찾지 못함: VI ID {target_vi_id}, 키워드 {target_keyword}")




if __name__ == "__main__":
    account = {"user_id": "sstrade251016", "user_pw": "a2345"}

    # 1. 수정을 위해 파일을 로드 (VBA 유지)
    main_wb = load_workbook(EXCEL_PATH, keep_vba=True)
    main_ws = main_wb['데이터']

    # 2. 날짜 열 동기화 (1/1 ~ 오늘) 및 검색 키워드 추출
    sync_date_columns_until_today(main_ws, start_date_str="2026-01-01")
    keywords = get_keyword_from_xlsm()

    # 3. 날짜-열 인덱스 매핑 생성 (YYYY-MM-DD 형식으로 통일)
    date_info_map = {}
    for col in range(74, main_ws.max_column + 1):
        val = main_ws.cell(row=5, column=col).value

        if val is None:
            break

        # 1. 셀 값이 이미 datetime 객체인 경우 (엑셀 날짜 형식)
        if isinstance(val, datetime):
            d_str = val.strftime('%Y-%m-%d')
            date_info_map[d_str] = col

        # 2. 셀 값이 텍스트인 경우 (1/1, 1/2 등)
        else:
            val_str = str(val).strip()
            try:
                # '2026/1/1'과 같은 형식으로 변환 시도
                # 변환에 성공하면 '1/1' 형태의 날짜로 간주
                temp_dt = datetime.strptime(f"2026/{val_str}", "%Y/%m/%d")
                d_str = temp_dt.strftime('%Y-%m-%d')
                date_info_map[d_str] = col
            except ValueError:
                # '1/1' 형식이 아니면 (공란, 직전 등 모든 텍스트 포함) 즉시 탐색 중단
                print(f"날짜 형식이 아닌 열 발견 ({val_str}). 날짜 수집을 종료합니다.")
                break

    # 4. 드라이버 실행 및 로그인
    driver = create_driver(account["user_id"], headless=False)

    try:
        if login_success_check(driver, account):
            # 5. 각 키워드별 반복 검색
            for keyword in keywords:
                # 이 키워드에 대해 순위가 비어있는 날짜들만 선별
                missing_dates = get_missing_dates_for_keyword(main_ws, keyword, date_info_map)

                if not missing_dates:
                    print(f"[{keyword}] 이미 모든 날짜의 데이터가 존재합니다. 건너뜁니다.")
                    continue

                print(f"\n--- {keyword} 작업 시작 (누락 날짜: {len(missing_dates)}개) ---")
                search_keyword(driver, keyword)

                # 6. 한 번의 페이지 스캔으로 누락된 모든 날짜 데이터 추출
                all_results_map = extract_product_results(driver, missing_dates)

                # 7. 추출된 결과를 날짜별로 엑셀에 반영
                for date_text, product_results in all_results_map.items():
                    for product_id, product_rank in product_results:
                        update_excel_rank(main_ws, product_id, keyword, product_rank, date_text)

            # 8. 모든 작업 완료 후 최종 저장 (성능을 위해 마지막에 한 번만 수행)
            main_wb.save(EXCEL_PATH)
            print("\n전체 작업 완료 및 엑셀 저장 성공!")

    except Exception as e:
        print(f"실행 중 오류 발생: {e}")
    finally:
        main_wb.close()
        driver.quit()