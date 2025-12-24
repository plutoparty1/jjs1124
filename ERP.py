from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# Chrome을 디버그 모드로 실행 (터미널 또는 명령 프롬프트에서 아래 명령어 실행)
# chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfile"

# Python 코드
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)

# 이후 driver를 사용해 브라우저 제어
driver.get("https://example.com")
