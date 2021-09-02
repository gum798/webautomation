from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import pandas as pd
from tkinter import *
from tkinter import messagebox

root = Tk()
root.withdraw()

df = pd.read_excel("keta_sample.xlsx")

print(df.columns)
df.info()
print(df['본인의 대륙 선택'][0])

driver= webdriver.Chrome()
url = 'https://www.k-eta.go.kr/portal/apply/viewstep1.do'
driver.get(url)
driver.maximize_window()
action = ActionChains(driver)


#STEP 01
driver.find_element_by_css_selector("#etaCnttGbcd").click()
action.send_keys(df['본인의 대륙 선택'][0]).perform()
action.reset_actions()

driver.find_element_by_css_selector("#natCd").click()

select = Select(driver.find_element_by_css_selector('#natCd'))
select.select_by_visible_text(df['본인의 국적 선택'][0])

#항목 개인정보 수집 동의(필수)
agree1 = driver.find_element_by_css_selector("#dAgree1")
action.move_to_element(agree1).click().perform()
action.reset_actions()

#선택 항목 개인정보 수집 동의
agree2 = driver.find_element_by_css_selector("#agree2")
action.move_to_element(agree2).click().perform()
action.reset_actions()

#민감정보 동의(필수)
agree3 = driver.find_element_by_css_selector("#dAgree3")
action.move_to_element(agree3).click().perform()
action.reset_actions()

#고유식별정보 동의(필수)
agree4 = driver.find_element_by_css_selector("#dAgree4")
action.move_to_element(agree4).click().perform()
action.reset_actions()

#이용약관 동의(필수)
agree5 = driver.find_element_by_css_selector("#dAgree5")
action.move_to_element(agree5).click().perform()
action.reset_actions()


print("")

messagebox.showinfo("작업완료", "입력내용을 확인하세요.\n\"확인\"버튼을 누르면 창이 닫힙니다. ")
