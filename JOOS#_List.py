from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import os
import glob
from datetime import datetime
import pyautogui
import time

# 크롬 드라이버 경로 설정 (본인 환경에 맞게 변경 필요)
chrome_driver_path = r'C:\Users\Code\chromedriver\chromedriver.exe'

# 크롬 옵션 설정
options = Options()
options.add_experimental_option('detach', True)
options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\정산-PC\OneDrive - 플랜티엠\경영지원팀의 파일 - 플랜티엠_정산 1\#. 매출발행\주스샵전체리스트",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True
})
driver = webdriver.Chrome(options=options)

# 서비스 객체 생성
service = Service(executable_path=chrome_driver_path)

# 드라이버 생성
driver = webdriver.Chrome(service=service, options=options)

# URL 접속
target_url = "http://121.254.227.75:7080/uat/uia/egovLoginUsr.do"
driver.get(target_url)

# 입력 칸 찾기 (XPath 이용)
id_input = driver.find_element(By.XPATH, '//*[@id="id"]')

# 아이디 입력
id_input.send_keys("jjs1124")

# 비밀번호 입력
pw_input = driver.find_element(By.XPATH, '//*[@id="password"]')
pw_input.send_keys("wjdwltn1!!")

# 로그인
login_btn = driver.find_element(By.XPATH, '//*[@id="command"]/fieldset/div/input')
login_btn.click()

# 고객
customer_btn = driver.find_element(By.XPATH, '//*[@id="topnavi"]/ul/li[3]/a')
customer_btn.click()

# 매장
customer_btn = driver.find_element(By.XPATH, '//*[@id="nav"]/div[2]/ul/li[2]/a')
customer_btn.click()

# 특정 요소(XPath) 나타날 때까지 최대 10초 대기
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="listForm"]/div[4]/table/thead/tr/th[11]'))
)

# 드롭박스 요소 찾기
dropdown = driver.find_element(By.XPATH, '//*[@id="searchCondition2"]')

# Select 클래스로 감싸기
select_box = Select(dropdown)

# "서비스중" 옵션 선택
select_box.select_by_visible_text("서비스중")

# 엑셀저장
customer_btn = driver.find_element(By.XPATH, '//*[@id="search_third_ul"]/li/div/a[6]')
customer_btn.click()

download_dir = r"C:\Users\정산-PC\OneDrive - 플랜티엠\경영지원팀의 파일 - 플랜티엠_정산 1\#. 매출발행\주스샵전체리스트"
today_str = datetime.now().strftime("%y%m%d")  # 251126 형식
new_filename = f"전체리스트_{today_str}.xls"
new_filepath = os.path.join(download_dir, new_filename)

# 다운로드 버튼 클릭 직후 대기 (예시: 5초)
time.sleep(5)
pyautogui.press('tab')

# 최신 파일 찾기: 다운로드 폴더 내 가장 최근 파일
list_of_files = glob.glob(os.path.join(download_dir, "*"))
latest_file = max(list_of_files, key=os.path.getctime)

# 원하는 이름으로 파일명 변경
os.rename(latest_file, new_filepath)
print(f"{latest_file} => {new_filepath} 파일명 변경 완료")

# 드롭박스에서 "매장탈퇴" 옵션 선택
select_box.select_by_visible_text("매장탈퇴")

# 엑셀저장 버튼 클릭 (이전과 동일)
customer_btn = driver.find_element(By.XPATH, '//*[@id="search_third_ul"]/li/div/a[6]')
customer_btn.click()

# 다운로드 후 대기
time.sleep(5)
pyautogui.press('tab')  # 필요에 따라 위치 유지

# 파일명 날짜(YYMMDD) 생성
today_str = datetime.now().strftime("%y%m%d")  # 예: 251126
new_filename = f"탈퇴리스트_{today_str}.xls"
new_filepath = os.path.join(download_dir, new_filename)

# 최신 파일 찾기: 다운로드 폴더 내 가장 최근 파일
list_of_files = glob.glob(os.path.join(download_dir, "*"))
latest_file = max(list_of_files, key=os.path.getctime)

# 파일명 변경
os.rename(latest_file, new_filepath)
print(f"{latest_file} => {new_filepath} 파일명 변경 완료")

# 모든 작업 완료 후 브라우저 종료
driver.quit()