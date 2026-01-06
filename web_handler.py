import time
import random
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from config import PROFILE_ROOT_DIR, TOP_ADS_URL


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