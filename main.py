import logging
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


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_ROOT_DIR = os.path.join(BASE_DIR, "Chrome_profile")

TOP_ADS_URL = "https://top.re.kr/ads"


def get_chrome_options(user_id: str, headless: bool = False, detach: bool = False) -> Options:
    options = Options()

    user_data_path = os.path.join(PROFILE_ROOT_DIR, user_id)

    # 공통
    options.add_argument(f"--user-data-dir={user_data_path}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # 기타 옵션
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

    return options


def create_driver(user_id, headless: bool = False) -> webdriver.Chrome:
    # 프로필 세팅용은 detach 필요 X
    options = get_chrome_options(user_id=user_id, headless=headless, detach=False)
    service = Service(ChromeDriverManager().install())

    try:
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_window_size(1300, 900)
        return driver
    except Exception as e:
        raise


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
        id_ele.click()
        type_like_human(id_ele, {account["user_id"]})
        time.sleep(1)

        # PW 한 글자씩 타이핑
        pw_ele.click()
        type_like_human(pw_ele, {account["user_pw"]})
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

        search_button.click()
        time.sleep(2)
        print(f"'{keyword}' 검색 완료")

    except Exception as e:
        print(f"키워드 검색 중 오류 발생: {e}")


EXCEL_PATH = os.path.join(BASE_DIR, 'TOP★점프_트래픽관리.xlsm')
def get_keyword_from_xlsm():
    if not os.path.exists(EXCEL_PATH):
        print(f"파일을 찾을 수 없습니다: {EXCEL_PATH}")
        return []

    # 1. 엑셀 로드
    # keep_vba=True: 매크로 유지
    # data_only=True: 수식이 아닌 '텍스트 결과값'만 가져옴
    wb = load_workbook(EXCEL_PATH, keep_vba=True, data_only=True)
    ws = wb['데이터']

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

    wb.close()
    return keywords


# date = '2026-01-01'인 경우만 처리
def extract_product_results(driver, target_date: str, timeout: int = 10):
    wait = WebDriverWait(driver, timeout)
    product_results = []

    try:
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))

        if not rows or (len(rows) == 1 and "슬롯 정보가 없습니다" in rows[0].text):
            print("조회 결과 없음")
            return []

        is_date_found = False  # 날짜 일치하는 행을 한 번이라도 만났는지 확인

        for row in rows:
            try:
                start_date_ele = row.find_element(By.XPATH, "./td[12]")
                start_date = start_date_ele.text.strip()

                if start_date == target_date:
                    is_date_found = True  # 일치하는 날짜 찾음
                    rank_text = row.find_element(By.XPATH, "./td[9]").text.strip()
                    rank_number = rank_text
                    if "순위밖" in rank_text:
                        rank_number = "순위밖"
                    elif "위" in rank_text:
                        # "위" 앞에 숫자가 있는 경우만 추출
                        rank_number = rank_text.split('위')[0].strip()

                    url = row.find_element(By.XPATH, "./td[8]//a").get_attribute("href")
                    last_id = url.split("=")[-1]

                    product_results.append((last_id, rank_number))
                    print(f"시작일: {start_date}, VI ID: {last_id}, 순위: {rank_number}")

                else:
                    print(f"{target_date}가 아니므로 작업 중단")
                    break  # 더이상 다음 row 보지 않고 for문 탈출 (내림차순이므로)

            except Exception as e:
                continue

        # 날짜는 찾았으나 결과가 비었을 때만 경고
        if not product_results:
            if is_date_found:
                print(f"'{target_date}' 행은 찾았으나, 내부 데이터(VI ID/순위) 추출에 실패했습니다.")

    except Exception as e:
        print(f"테이블 로딩 중 오류 발생: {e}")

    return product_results


def update_excel_rank(target_vi_id, target_keyword, rank_value):
    if not os.path.exists(EXCEL_PATH):
        print("파일을 찾을 수 없습니다.")
        return

    # keep_vba=True 옵션으로 매크로 보존
    # 값을 입력해야 하므로 data_only=False(기본값)
    wb = load_workbook(EXCEL_PATH, keep_vba=True)
    ws = wb['데이터']

    # 열 번호 설정
    COL_VI_ID = 6  # F열
    COL_KEYWORD = 10  # J열
    COL_TARGET = 74  # BV열

    found = False
    # 7행부터 마지막 행까지 탐색
    for row in range(7, ws.max_row + 1):
        # 엑셀의 VI ID가 숫자형일 수 있으므로 문자열로 변환하여 비교
        vi_id = str(ws.cell(row=row, column=COL_VI_ID).value or "").strip()
        keyword = str(ws.cell(row=row, column=COL_KEYWORD).value or "").strip()

        # 두 조건이 일치하는 행 찾기
        if vi_id == str(target_vi_id) and keyword == target_keyword:
            # BV열(74번)에 값 입력
            ws.cell(row=row, column=COL_TARGET).value = rank_value
            print(f"{row}행에 순위 값 '{rank_value}' 입력")
            found = True
            break  # 찾았으므로 루프 종료 (1개 밖에 없는 게 맞는지 확인 필요)

    if not found:
        print(f"조건에 맞는 행을 찾지 못함: VI ID {target_vi_id}, 키워드 {target_keyword}")



if __name__ == "__main__":
    account = {"user_id": "sstrade251016", "user_pw": "a2345"}
    user_id = account["user_id"]

    wb = load_workbook(EXCEL_PATH, keep_vba=True, data_only=True)

    driver = create_driver(user_id, headless=False)

    login_success = login_success_check(driver, account)

    keywords = get_keyword_from_xlsm()
    for keyword in keywords:
        search_keyword(driver, keyword)

        product_results = extract_product_results(driver, '2026-01-01')

        for product_result in product_results:
            product_id = product_result[0]
            product_rank = product_result[1]

            update_excel_rank(product_id, keyword, product_rank)

    wb.save(EXCEL_PATH)
    print("파일이 저장되었습니다.")

    wb.close()
