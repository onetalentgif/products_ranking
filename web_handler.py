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
import shutil  # 폴더 삭제를 위해



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
        pw_ele.clear()
        time.sleep(0.5)

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

        time.sleep(1.5)
        print(f"'{keyword}' 검색 완료")

    except Exception as e:
        print(f"키워드 검색 중 오류 발생: {e}")


# target_dates = ['2026-01-07', '2026-01-08'] (텍스트 형식, 반드시 날짜 순서 유지해야 함, 오늘 날짜까지만!)
def extract_product_results(driver, target_dates: list, timeout: int = 10):
    wait = WebDriverWait(driver, timeout)

    # 타겟 날짜 텍스트를 datetime 객체로 변환 (리스트)
    target_datetimes = [datetime.strptime(target_date, '%Y-%m-%d') for target_date in target_dates]
    min_target = min(target_datetimes)
    max_target = max(target_datetimes)

    product_results = {target_datetime: [] for target_datetime in target_datetimes}  # product_results = { datetime(2026, 1, 7): [], datetime(2026, 1, 8): [] }

    try:
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))

        # '조회 결과 없음' 문구가 있는 경우 처리
        if len(rows) == 0 or (len(rows) == 1 and "정보가 없습니다" in rows[0].text):
            print("조회 결과 없음 (표시된 데이터가 없습니다)")
            return product_results

        for row in rows:
            try:
                start_date_text = row.find_element(By.XPATH, "./td[12]").text.strip()[:10] # (아이콘 제거)
                start_date = datetime.strptime(start_date_text, '%Y-%m-%d')

                end_date_text = row.find_element(By.XPATH, "./td[13]").text.strip()[:10] # (아이콘 제거)
                end_date = datetime.strptime(end_date_text, '%Y-%m-%d')

                # 종료일이 지났으면(타겟 날짜에 해당하는 기간이 없으면) 중단
                if end_date < min_target:
                    print(f"종료일({end_date_text})이 지났으므로 탐색 종료")
                    break   # break 로직은 데이터가 날짜순일 때만 유효

                # 시작일이 아직 안왔으면 다음 행으로 이동
                if start_date > max_target:
                    continue

                # 모든 타겟 날짜에 대해 매칭 확인
                for i, target_datetime in enumerate(target_datetimes):
                    if start_date <= target_datetime <= end_date:

                        row_keyword = row.find_element(By.XPATH, "./td[6]").text.strip()

                        url = row.find_element(By.XPATH, "./td[8]//a").get_attribute("href")
                        product_id = url.split("=")[-1]

                        # 해당 날짜의 키워드 & 상품 번호 중복 체크
                        if any(item[0] == row_keyword and item[1] == product_id for item in product_results[target_datetime]):
                            # print(f"중복 데이터 건너뜀: {row_keyword} | {product_id}")
                            continue

                        rank_text = row.find_element(By.XPATH, "./td[9]").text.strip()
                        rank_number = ""
                        if "순위밖" not in rank_text and "위" in rank_text:
                            rank_number = rank_text.split('위')[0].strip()

                        product_results[target_datetimes[i]].append((row_keyword, product_id, rank_number))
                        print(f"매칭 발견: {target_datetime} | 키워드: {row_keyword} | ID: {product_id} | 순위: {rank_number}")

            except Exception as e:
                # 개별 행 파싱 실패 시 다음 행으로 진행
                print(f"행 처리 중 오류: {e}")
                continue

    except Exception as e:
        print(f"테이블 처리 중 오류 발생: {e}")

    return product_results



def delete_chrome_cache(user_id):
    target_profile_path = os.path.join(PROFILE_ROOT_DIR, user_id)

    if not os.path.exists(target_profile_path):
        print(f"SKIP: 프로필 폴더가 존재하지 않습니다. ({target_profile_path})")
        return

    print(f"\n[Cache Cleanup] Start: {target_profile_path}")

    # 삭제할 하위 폴더 목록 (Default 내부 및 공통 캐시)
    target_folders = [
        os.path.join("Default", "Cache"),
        os.path.join("Default", "Code Cache"),
        os.path.join("Default", "GPUCache"),
        "component_crx_cache",
        "GrShaderCache",
        "optimization_guide_model_store"
    ]

    for sub_folder in target_folders:
        full_folder_path = os.path.join(target_profile_path, sub_folder)

        # 폴더가 실제로 있을 때만 삭제 시도
        if os.path.exists(full_folder_path):
            try:
                shutil.rmtree(full_folder_path)
                print(f"SUCCESS: Deleted {sub_folder}")
            except Exception as e:
                # 보통 '액세스 거부' 에러가 많음 (크롬이 켜져 있을 때)
                print(f"FAIL: {sub_folder} / {e}")
        else:
            print(f"SKIP: {sub_folder} (폴더 없음)")

    print("[Cache Cleanup] Finished.\n")



from excel_handler import get_keyword_from_xlsm

if __name__ == "__main__":
    account = {"user_id": "sstrade251016", "user_pw": "a2345"}

    target_dates = ['2026-01-06', '2026-01-07'] # 반드시 날짜 순서 유지 (오름차순), 텍스트 형식

    # 작업 시작 전 불필요한 캐시 삭제
    delete_chrome_cache(account["user_id"])

    # 검색할 키워드 목록 가져오기
    keywords = get_keyword_from_xlsm()

    # 드라이버 실행
    driver = create_driver(account["user_id"], headless=False)

    try:
        # 로그인 상태 확인
        if login_success_check(driver, account):
            # 각 키워드별 반복 검색
            for keyword in keywords:
                search_keyword(driver, keyword)

                # 웹 페이지에서 키워드, VI ID, 순위 추출 (딕셔너리 형태)
                product_results = extract_product_results(driver, target_dates)

                # 추출된 결과(키워드, ID, 순위) 출력
                # {datetime(2026, 1, 7): [('keyword1', 'ID_1', '3'), ('keyword2', 'ID_1', '10')], datetime(2026, 1, 8): [('keyword1', 'ID_1', '3')]}
                for target_date, items in product_results.items():
                    date_str = target_date.strftime('%Y-%m-%d')  # 출력용 날짜 변환

                    # product_results가 비었을 때
                    if not items:
                        print(f" {keyword} [{date_str}] 검색된 상품이 없습니다.")
                        continue

                    # 각 날짜에 포함된 상품 리스트를 순회 (row_keyword, product_id, rank_number)
                    for item in items:
                        product_keyword = item[0]   # 튜플의 첫 번째: 행 키워드
                        product_id = item[1]  # 튜플의 두 번째: ID
                        product_rank = item[2]  # 튜플의 세 번째: 순위

                        print(f" [{date_str}] 키워드: {product_keyword} | ID: {product_id} | 순위: {product_rank}")

    except Exception as e:
        print(f"실행 중 오류 발생: {e}")
    finally:
        driver.quit()
        print("드라이버가 종료되었습니다.")
