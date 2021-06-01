import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from datetime import date, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

today = date.today()
ftoday = date.today() - timedelta(1)
yesterdayDate = ftoday.strftime('%Y-%m-%d')
#해당 달의 첫째날 구하기
first_day = ftoday.replace(day=1)
first_dayDate = first_day.strftime('%Y-%m-%d')
#전달의 마지막 날 구하기
last_day_month_ago = first_day - timedelta(days=1)
last_day_month_agoDate = last_day_month_ago.strftime('%Y-%m-%d')
#전달의 첫째 날 구하기
first_day_month_ago = last_day_month_ago.replace(day=1)
first_day_month_agoDate = first_day_month_ago.strftime('%Y-%m-%d')
# -------------------------------------------------------------------------------
# 전산 로그인
driver = webdriver.Chrome()
driver.get("https://igms.golfzonretail.local/base/login/loginForm.do")
driver.maximize_window()
elem = driver.find_element_by_id('j_username')
elem.send_keys("skang35")
elem = driver.find_element_by_id('j_password')
elem.send_keys("golfzon1!")
elem.send_keys(Keys.RETURN)
# -------------------------------------------------------------------------------
# 입고/반출현황
time.sleep(1)
elem = driver.find_element_by_link_text("물류")
elem.click()
elem = driver.find_element_by_id("13f3bc56a9e3")
elem.click()
elem = driver.find_element_by_id("13f3bccfd7e9")
elem.click()
elem = driver.find_element_by_link_text("입고/반출현황")
elem.click()

time.sleep(1)
# 검색조건으로 iframe 이동
iframes = driver.find_elements_by_tag_name('iframe')
driver.switch_to.frame(iframes[1])

# 기간 입력
elem = driver.find_element_by_id("s_startCnfmDate")
elem.clear()
elem.send_keys(first_dayDate)
elem = driver.find_element_by_id("s_endCnfmDate")
elem.clear()
elem.send_keys(yesterdayDate)

# 분류 보임 클릭
driver.find_element_by_xpath('//*[@id="formSearch"]/div/table/tbody/tr[5]/td[3]/label[1]/input').click()
elem = driver.find_element_by_id("btnSearch")
elem.click()
time.sleep(6)
elem = driver.find_element_by_id("basicExcel")
elem.click()
time.sleep(1)
# -------------------------------------------------------------------------------
# 통합수불현황
driver.switch_to.parent_frame()
elem = driver.find_element_by_id("13f3bc7b2087")
elem.click()
elem = driver.find_element_by_id("13f3bda6dff3")
elem.click()
elem = driver.find_element_by_link_text("통합수불현황")
elem.click()
time.sleep(2)

iframes = driver.find_elements_by_tag_name('iframe')
driver.switch_to.frame(iframes[2])

elem = driver.find_element_by_id("startDate")
elem.clear()
elem.send_keys("202105")

# 통합수불 분유 보임
driver.find_element_by_xpath('//*[@id="subContentBox"]/div[1]/div[1]/div[1]/table/tbody/tr[3]/td[3]/label[1]/input').click()
elem = driver.find_element_by_id("btnSearch")
elem.click()
time.sleep(14)
elem = driver.find_element_by_id("basicExcel")
elem.click()
# -------------------------------------------------------------------------------
# BEST판매
driver.switch_to.parent_frame()
elem = driver.find_element_by_link_text("분석/통계")
elem.click()
elem = driver.find_element_by_id("157fe5c1e8c2")
elem.click()
elem = driver.find_element_by_id("165310f9a9e5")
elem.click()
elem = driver.find_element_by_link_text("BEST상품 판매현황(상품)")
elem.click()

time.sleep(1)
# 검색조건으로 iframe 이동
iframes = driver.find_elements_by_tag_name('iframe')
driver.switch_to.frame(iframes[3])

# 날짜 0차 셋팅 (전일)
elem = driver.find_element_by_id("BGN_DATE")
elem.clear()
elem.send_keys(yesterdayDate)
elem = driver.find_element_by_id("CL_DATE")
elem.clear()
elem.send_keys(yesterdayDate)

elem = driver.find_element_by_id("btnSearch")
elem.click()
time.sleep(11)
elem = driver.find_element_by_id("btnExcel")
elem.click()
time.sleep(1)
da = Alert(driver)
da.accept()
# -------------------------------------------------------------------------------
# 날짜 1차 셋팅 (1일 ~ 전일)
elem = driver.find_element_by_id("BGN_DATE")
elem.clear()
elem.send_keys("2021-05-01")
elem = driver.find_element_by_id("CL_DATE")
elem.clear()
elem.send_keys(yesterdayDate)

elem = driver.find_element_by_id("btnSearch")
elem.click()
time.sleep(35)
elem = driver.find_element_by_id("btnExcel")
elem.click()
time.sleep(1)
da = Alert(driver)
da.accept()
# -------------------------------------------------------------------------------
# 날짜 2차 셋팅(전월)
elem = driver.find_element_by_id("BGN_DATE")
elem.clear()
elem.send_keys(first_day_month_agoDate)
elem = driver.find_element_by_id("CL_DATE")
elem.clear()
elem.send_keys(last_day_month_agoDate)

elem = driver.find_element_by_id("btnSearch")
elem.click()
time.sleep(32)
elem = driver.find_element_by_id("btnExcel")
elem.click()
time.sleep(1)
da = Alert(driver)
da.accept()
# -------------------------------------------------------------------------------
# 상품코드
driver.switch_to.parent_frame()
elem = driver.find_element_by_link_text("기준정보")
elem.click()
elem = driver.find_element_by_id("1638d07f5b73")
elem.click()
elem = driver.find_element_by_xpath('//*[@id="m3_1638d07f5b73"]/li[5]')
elem.click()

time.sleep(2)
iframes = driver.find_elements_by_tag_name('iframe')
driver.switch_to.frame(iframes[4])
elem = driver.find_element_by_id("prdInq_btnSearch")
elem.click()
time.sleep(20)
elem = driver.find_element_by_id("prdInq_btnExcelDownLoad")
elem.click()
time.sleep(11)
da = Alert(driver)
da.accept()

driver.quit()