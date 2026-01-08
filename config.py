import os
import sys


def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# 경로 설정
BASE_DIR = get_base_dir()
# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_ROOT_DIR = os.path.join(BASE_DIR, "Chrome_profile")
EXCEL_PATH = os.path.join(BASE_DIR, 'TOP★점프_트래픽관리.xlsm')

# 웹 설정
TOP_ADS_URL = "https://top.re.kr/ads"
TARGET_START_DATE = "2026-01-01"

# 계정 정보
ACCOUNT = {"user_id": "sstrade251016", "user_pw": "a2345"}

# 최대 검색할 페이지 수 설정
MAX_PAGES = 4