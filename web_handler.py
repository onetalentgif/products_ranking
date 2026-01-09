import time
import random
import os
import shutil
from typing import List, Dict, Optional, Tuple, Any

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.support.ui import Select
from datetime import datetime
from config import PROFILE_ROOT_DIR, TOP_ADS_URL


class Locators:
    """XPaths and CSS Selectors for the web handler."""
    # Login Page
    ID_INPUT = (By.XPATH, '//label[contains(text(), "아이디")]/following-sibling::input')
    PW_INPUT = (By.XPATH, '//label[contains(text(), "비밀번호")]/following-sibling::input')
    REMEMBER_CHECKBOX = (By.ID, "remember")
    LOGIN_BTN = (By.XPATH, '//button[contains(text(), "로그인")]')
    LOGOUT_BTN = (By.XPATH, '//button[contains(text(), "로그아웃")]')

    # Main Page / Search
    SEARCH_INPUT = (By.XPATH, "//input[contains(@placeholder, '슬롯번호, 아이디, 키워드')]")
    SEARCH_BTN = (By.XPATH, "//button[contains(text(), '검색')]")
    NORMAL_TAB = (By.XPATH, "//div[contains(@class, 'cursor-pointer')]//p[contains(text(), '정상')]")
    
    # Pagination & View Settings
    VIEW_SELECT = (By.XPATH, "//select[option[@value='1000']]")
    NEXT_PAGE_BTN = (By.XPATH, "//button[@title='다음 페이지']")
    
    # Table Results
    TABLE_ROWS = (By.XPATH, "//tbody/tr")
    ROW_START_DATE = (By.XPATH, "./td[12]")
    ROW_END_DATE = (By.XPATH, "./td[13]")
    ROW_KEYWORD = (By.XPATH, "./td[6]")
    ROW_LINK = (By.XPATH, "./td[8]//a")
    ROW_RANK_TEXT = (By.XPATH, "./td[9]")


def create_driver(user_id: str, headless: bool = False) -> webdriver.Chrome:
    """Initialize and return a Chrome WebDriver with specific options."""
    options = Options()
    user_data_path = os.path.join(PROFILE_ROOT_DIR, user_id)
    options.add_argument(f"--user-data-dir={user_data_path}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    if headless:
        options.add_argument("--headless=new")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(1300, 900)
    return driver


def type_like_human(element: WebElement, text: str) -> None:
    """Simulate human typing by introducing random delays between keystrokes."""
    for char in text:
        element.send_keys(char)
        delay = random.uniform(0.1, 0.4) # Slightly faster than before but still human-like
        time.sleep(delay)


def login_top_with_send_keys(driver: webdriver.Chrome, account: Dict[str, str], debug: bool = True) -> bool:
    """Perform the login sequence using human-like typing."""
    try:
        if debug:
            print(f"로그인 시도: {account['user_id']}")

        driver.get(TOP_ADS_URL)
        wait = WebDriverWait(driver, 10)

        id_ele = wait.until(EC.element_to_be_clickable(Locators.ID_INPUT))
        pw_ele = driver.find_element(*Locators.PW_INPUT)

        # ID Typing
        id_ele.clear()
        time.sleep(0.5)
        id_ele.click()
        type_like_human(id_ele, account["user_id"])
        time.sleep(0.5)

        # PW Typing
        pw_ele.clear()
        time.sleep(0.5)
        pw_ele.click()
        type_like_human(pw_ele, account["user_pw"])
        time.sleep(0.5)

        # Remember Me Checkbox
        keep_ele = driver.find_element(*Locators.REMEMBER_CHECKBOX)
        keep_ele.click()
        time.sleep(0.5)

        # Login Button
        login_btn = driver.find_element(*Locators.LOGIN_BTN)
        login_btn.click()
        
        # Verify login success by waiting for URL change or Logout button
        # (This will be handled by the caller or immediate check)
        time.sleep(2) 

        return True

    except Exception as e:
        if debug:
            print(f"[LOGIN] 로그인 중 예외: {e}")
        return False


def is_top_logged_in(driver: webdriver.Chrome, timeout: int = 3) -> bool:
    """Check if the user is currently logged in."""
    driver.get(TOP_ADS_URL)
    wait = WebDriverWait(driver, timeout)

    # 1) Check for Logout button
    try:
        logout_btn = wait.until(EC.presence_of_element_located(Locators.LOGOUT_BTN))
        if logout_btn.is_displayed():
            # print("로그아웃 버튼이 보여서, 로그인 상태로 판단합니다.")
            return True
    except TimeoutException:
        pass
    except Exception as e:
        print("로그아웃 버튼 확인 중 예외:", e)

    # 2) Check for Login button
    try:
        login_btn = driver.find_element(*Locators.LOGIN_BTN)
        if login_btn.is_displayed():
            # print("로그인 버튼이 보여서, 아직 로그인 안 된 상태로 판단합니다.")
            return False
    except NoSuchElementException:
        pass
    except Exception as e:
        print("로그인 버튼 확인 중 예외:", e)

    # If neither found immediately, assume logged out to be safe
    print("로그인/로그아웃 버튼을 찾지 못했습니다. 로그아웃 상태로 간주합니다.")
    return False


def login_success_check(driver: webdriver.Chrome, account: Dict[str, str]) -> bool:
    """Ensure the user is logged in, retrying if necessary."""
    user_id = account['user_id']

    try:
        if is_top_logged_in(driver, 3):
            print(f"[{user_id}] 이미 로그인되어 있습니다.")
            return True
    except Exception as e:
        print(f"로그인 상태 확인 중 에러: {e}")

    print(f"[{user_id}] 로그인 세션이 없습니다. 로그인을 시도합니다.")
    MAX_RETRIES = 3

    for attempt in range(1, MAX_RETRIES + 1):
        print(f"로그인 시도 {attempt}/{MAX_RETRIES}회 진행 중...")

        try:
            if login_top_with_send_keys(driver, account):
                if is_top_logged_in(driver, 5):
                    print(f"[{user_id}] 로그인 성공! 작업을 시작합니다.")
                    return True
        except Exception as e:
            print(f"[{user_id}] 시도 중 에러 발생: {e}")

        if attempt < MAX_RETRIES:
            print(f"[{user_id}] 로그인 실패. {attempt+1}회차 재시도를 위해 대기합니다.")
            time.sleep(3)
        else:
            print(f"[{user_id}] 모든 재시도 횟수를 소진했습니다.")

    print(f"[{user_id}] 최종적으로 로그인에 실패하였습니다.")
    return False


def search_keyword(driver: webdriver.Chrome, keyword: str, timeout: int = 10) -> bool:
    """Enter keyword and click search."""
    try:
        driver.get(TOP_ADS_URL)
        wait = WebDriverWait(driver, timeout)

        search_input = wait.until(EC.element_to_be_clickable(Locators.SEARCH_INPUT))
        search_input.clear()
        search_input.send_keys(keyword)

        search_button = wait.until(EC.element_to_be_clickable(Locators.SEARCH_BTN))
        # Use execute_script for reliable clicking
        driver.execute_script("arguments[0].click();", search_button)

        # Wait for the results to potentially load or the 'Processing' overlay to vanish
        # Since we don't have a reliable "loaded" indicator, a short sleep is often safest for resets
        time.sleep(1.0) 
        print(f"'{keyword}' 검색 완료")
        return True

    except Exception as e:
        print(f"키워드 검색 중 오류 발생: {e}")
        return False


def extract_product_results(driver: webdriver.Chrome, target_dates: List[str], timeout: int = 10) -> Dict[datetime, List[Tuple[str, str, str]]]:
    """
    Extract ranking results for the given target dates from the current page(s).
    Returns a dictionary mapping datetime objects to a list of (keyword, product_id, rank).
    """
    wait = WebDriverWait(driver, timeout)
    set_page_view_to_1000(driver)

    # Convert date strings to datetime objects once
    target_datetimes = [datetime.strptime(td, '%Y-%m-%d') for td in target_dates]
    min_target = min(target_datetimes)
    max_target = max(target_datetimes)

    # Result structure: { datetime: [(keyword, id, rank), ...], ... }
    product_results = {dt: [] for dt in target_datetimes}

    page_num = 1
    while True:
        print(f"  - {page_num}페이지 데이터 추출 중...")
        try:
            # Wait for table rows to appear
            rows = wait.until(EC.presence_of_all_elements_located(Locators.TABLE_ROWS))

            # Check for "Empty" message
            if not rows or (len(rows) == 1 and "정보가 없습니다" in rows[0].text):
                print("조회 결과 없음 (표시된 데이터가 없습니다)")
                break

            for row in rows:
                try:
                    # Extract dates
                    start_date_text = row.find_element(*Locators.ROW_START_DATE).text.strip()[:10]
                    end_date_text = row.find_element(*Locators.ROW_END_DATE).text.strip()[:10]
                    
                    try:
                        start_date = datetime.strptime(start_date_text, '%Y-%m-%d')
                        end_date = datetime.strptime(end_date_text, '%Y-%m-%d')
                    except ValueError:
                        # Date parsing failed, skip row
                        continue

                    # Optimization: Skip if date range doesn't overlap with ANY of our targets roughly
                    # However, since rows might not be sorted, we shouldn't break the loop completely
                    # unless we are sure of the sort order.
                    # We continue row processing only if it might contain relevant data.

                    if end_date < min_target:
                        continue
                    if start_date > max_target:
                        continue

                    # Check which specific target dates this row covers
                    start_date_ts = start_date.timestamp()
                    end_date_ts = end_date.timestamp()

                    # Extract details only if needed
                    row_keyword = ""
                    product_id = ""
                    rank_number = ""
                    
                    details_extracted = False

                    for target_dt in target_datetimes:
                        target_ts = target_dt.timestamp()
                        
                        if start_date_ts <= target_ts <= end_date_ts:
                            # Lazy extraction
                            if not details_extracted:
                                row_keyword = row.find_element(*Locators.ROW_KEYWORD).text.strip()
                                url = row.find_element(*Locators.ROW_LINK).get_attribute("href")
                                product_id = url.split("=")[-1]
                                
                                rank_text = row.find_element(*Locators.ROW_RANK_TEXT).text.strip()
                                if "순위밖" in rank_text:
                                    rank_number = "X"
                                elif "위" in rank_text:
                                    rank_number = rank_text.split('위')[0].strip()
                                else:
                                    rank_number = ""
                                details_extracted = True

                            # Check for duplicates associated with this specific date
                            # (keyword, id) pair logic check
                            is_duplicate = False
                            for item in product_results[target_dt]:
                                if item[0] == row_keyword and item[1] == product_id:
                                    is_duplicate = True
                                    break
                            
                            if not is_duplicate:
                                product_results[target_dt].append((row_keyword, product_id, rank_number))
                                print(f"매칭 발견: {target_dt.strftime('%Y-%m-%d')} | 키워드: {row_keyword} | ID: {product_id} | 순위: {rank_number}")

                except StaleElementReferenceException:
                    # Row updated during iteration, skip safely
                    continue
                except Exception as e:
                    print(f"행 처리 중 오류: {e}")
                    continue

        except Exception as e:
            print(f"테이블 처리 중 오류 발생: {e}")
            break

        # Pagination
        if go_to_next_page(driver):
            page_num += 1
            # Give a small buffer for table refresh
            time.sleep(1.0) 
        else:
            print(f"  >>> 모든 페이지({page_num}p) 수집 완료.")
            break

    return product_results


def set_page_view_to_1000(driver: webdriver.Chrome, timeout: int = 10) -> None:
    """Ensure the page is showing 1000 results per page."""
    try:
        wait = WebDriverWait(driver, timeout)
        
        # Locate the select element containing option '1000'
        select_element = wait.until(EC.presence_of_element_located(Locators.VIEW_SELECT))
        select = Select(select_element)

        if select.first_selected_option.get_attribute("value") == "1000":
            # print("이미 페이지 보기가 1000개로 설정되어 있습니다.")
            return

        select.select_by_value("1000")
        print("페이지 보기를 1000개로 설정했습니다.")
        
        # Determine if the table is refreshing.
        # Simple wait is often more robust than complex stale checks here.
        time.sleep(1.5)

    except Exception as e:
        print(f"1000개 보기 설정 중 오류 발생: {e}")


def go_to_next_page(driver: webdriver.Chrome, timeout: int = 5) -> bool:
    """Click the 'Next Page' button. Returns True if successful, False if disabled/not found."""
    try:
        wait = WebDriverWait(driver, timeout)
        next_btn = wait.until(EC.presence_of_element_located(Locators.NEXT_PAGE_BTN))

        if next_btn.get_attribute("disabled") is not None:
            return False

        driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
        time.sleep(0.3) # Short animation wait
        driver.execute_script("arguments[0].click();", next_btn)
        return True
    except Exception:
        return False


def extract_normal_products(driver: webdriver.Chrome, target_dates: List[str]) -> Optional[Dict[datetime, List[Tuple[str, str, str]]]]:
    """Filter by 'Normal' status and extract data."""
    wait = WebDriverWait(driver, 10)

    try:
        # Click 'Normal' (정상)
        normal_btn = wait.until(EC.element_to_be_clickable(Locators.NORMAL_TAB))
        normal_btn.click()
        print("'정상' 필터 선택 완료")
        time.sleep(0.5)

        # Click Search
        search_btn = wait.until(EC.element_to_be_clickable(Locators.SEARCH_BTN))
        search_btn.click()
        print("검색 실행")
        time.sleep(1.5)

        # Extract (reuses the same logic)
        return extract_product_results(driver, target_dates)

    except Exception as e:
        print(f"전체 수집 과정 중 오류 발생: {e}")
        return None


def kill_chrome_processes() -> None:
    """Kill dangling Chrome processes."""
    print("기존 크롬 프로세스를 종료하는 중...")
    try:
        os.system("taskkill /f /im chrome.exe /t >nul 2>&1")
        os.system("taskkill /f /im chromedriver.exe /t >nul 2>&1")
        print("프로세스 종료 완료.")
        time.sleep(1)
    except Exception as e:
        print(f"프로세스 종료 중 오류(무시 가능): {e}")


def delete_chrome_cache(user_id: str) -> None:
    """Clear Chrome cache for the specific user profile."""
    target_profile_path = os.path.join(PROFILE_ROOT_DIR, user_id)

    if not os.path.exists(target_profile_path):
        return

    print(f"\n[Cache Cleanup] Start: {target_profile_path}")

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

        if os.path.exists(full_folder_path):
            try:
                shutil.rmtree(full_folder_path)
                # print(f"SUCCESS: Deleted {sub_folder}")
            except Exception as e:
                # Often access denied if Chrome is still closing; usually fine to ignore
                pass
                # print(f"FAIL: {sub_folder} / {e}")

    print("[Cache Cleanup] Finished.\n")


# For standalone testing
if __name__ == "__main__":
    test_account = {"user_id": "test_id", "user_pw": "test_pw"}
    test_dates = [datetime.now().strftime('%Y-%m-%d')]
    
    # Simple validation that the function signatures are correct
    print("Testing WebHandler imports...")
    try:
        # kill_chrome_processes()
        print("WebHandler module loaded successfully.")
    except Exception as e:
        print(f"Error loading module: {e}")
