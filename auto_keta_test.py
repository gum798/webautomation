import json
import os
import sys
import time
from ftplib import FTP
from tkinter import *
from tkinter import messagebox

import pywinauto
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

root = Tk()
root.withdraw()

if len(sys.argv)>1:
    messagebox.showinfo("파라미터", sys.argv[1])
    print("")
else:
    messagebox.showinfo("파라미터", "파라미터가 없습니다.")
    sys.exit()

params = json.loads(sys.argv[1].replace("keta://",""))


def download_file(ftp, dir, file):
    ftp.cwd("/")
    ftp.cwd(dir[0])
    ftp.cwd(dir[1])
    download_path = os.path.abspath("./tmp/")
    os.makedirs(download_path, exist_ok=True)
    with open(download_path + file, 'wb') as localfile:
        ftp.retrbinary('RETR ' + file, localfile.write)
        print('Download file is ' + file)
    abspath = os.path.abspath(download_path + file)
    return abspath

def path_parser(path):
    srcDir = []
    split_file = path.split("/")
    srcDir.append(split_file[3])
    srcDir.append(split_file[4])
    srcFileName = split_file[5]
    return srcDir, srcFileName


filepath = "/nas/keta/KOR001_2021/ERN202110000000000317/ERA202110000000000582.jpg"
filepath2 = "/nas/keta/KOR001_2021/ERN202110000000000317/ERA202110000000000581.jpg"



#FTP접속
ftp = FTP('121.78.145.131')
ftp.login("ftpketa","ftpketa!234")
print(ftp.retrlines('LIST'))
srcDir, srcFileName = path_parser(filepath)
IMG_PASSPORT = download_file(ftp, srcDir, srcFileName)
srcDir, srcFileName = path_parser(filepath2)
IMG_FACE = download_file(ftp, srcDir, srcFileName)

ftp.quit()

#크롬드라이버 로딩
driver= webdriver.Chrome()
url = 'https://www.k-eta.go.kr/portal/apply/viewstep1.do'
driver.get(url)
driver.maximize_window()
action = ActionChains(driver)

#STEP 01
#대륙선택------------------------
driver.find_element_by_css_selector("#etaCnttGbcd").click()
action.send_keys("AMERICA").perform()
action.reset_actions()
driver.find_element_by_css_selector("#natCd").click()
#대륙선택------------------------


#국적선택------------------------
select = Select(driver.find_element_by_css_selector('#natCd'))
select.select_by_value(params.get('passportPublishCountry'))
#국적선택------------------------

#신청불가------------------------
exit_flag = False
try:
#이미작성된내용-취소------------------------
    ok = driver.find_element_by_css_selector(".popBtn2")
    if ok != None:
        messagebox.showinfo("신청불가", "\"확인\"버튼을 누르면 종료됩니다.")
        exit_flag = True
except:
    print("")
if exit_flag:
    sys.exit()
#신청불가------------------------


#필수개인정보수집동의------------------------
agree1 = driver.find_element_by_css_selector("#dAgree1")
action.move_to_element(agree1).click().perform()
action.reset_actions()
#필수개인정보수집동의------------------------

#선택개인정보수집동의------------------------
agree2 = driver.find_element_by_css_selector("#dAgree2")
action.move_to_element(agree2).click().perform()
action.reset_actions()
#선택개인정보수집동의------------------------

#민감정보처리동의------------------------
agree3 = driver.find_element_by_css_selector("#dAgree3")
action.move_to_element(agree3).click().perform()
action.reset_actions()
#민감정보처리동의------------------------

#고유식별정보처리동의------------------------
agree4 = driver.find_element_by_css_selector("#dAgree4")
action.move_to_element(agree4).click().perform()
action.reset_actions()
#고유식별정보처리동의------------------------

#약관동의------------------------
agree5 = driver.find_element_by_css_selector("#dAgree5")
action.move_to_element(agree5).click().perform()
action.reset_actions()
#약관동의------------------------

#다음------------------------
driver.find_element_by_css_selector(".btnBasic2.popupCheck").click()
#다음------------------------

#STEP 02

#여권번호------------------------
driver.find_element_by_css_selector("#psno").send_keys(params.get('passportNo'))
#여권번호------------------------

#이메일주소------------------------
driver.find_element_by_css_selector("#emlAddr").send_keys(params.get('email'))
#이메일주소------------------------

#확인------------------------
driver.find_element_by_css_selector(".btnInput1").click()
#확인------------------------

time.sleep(0.5)

#STEP 03
try:
#이미작성된내용-취소------------------------
    cancel = driver.find_element_by_css_selector(".popBtn1")
    if cancel != None:
        driver.find_element_by_css_selector(".popBtn1").click()
#이미작성된내용-취소------------------------

#이미작성된내용-확인------------------------
    # load = driver.find_element_by_css_selector(".popBtn2")
    # if load != None:
    #     driver.find_element_by_css_selector(".popBtn2").click()
    #     messagebox.showinfo("임시저장 불러오기완료", "입력내용을 확인하세요.\n\"확인\"버튼을 누르면 창이 닫힙니다. ")
    #     exit()
#이미작성된내용-확인------------------------
except:
    print("")

# 여권사진 업로드
# driver.find_element_by_css_selector(".irx_filetree_list.irx_listGrid").send_keys("C:\\Users\\jhseo\\Desktop\\여권.png")
#파일추가------------------------
driver.find_element_by_css_selector("#btnAddFile").click()
time.sleep(3)
#파일추가------------------------

#파일열기------------------------
app = pywinauto.Application().connect(title="열기", class_name="#32770")
dlg = app.window(title_re="열기", class_name="#32770")
dlg.Edit1.SetEditText(IMG_PASSPORT)
dlg.Button1.click()
#파일열기------------------------
time.sleep(3)

#성별------------------------
if params.get('gender') == "M" :
    #남자------------------------
    btn = driver.find_element_by_xpath('//input[@id = "sdCd_M"]')
    driver.execute_script("arguments[0].click();", btn)
    #남자------------------------
else:
    #여자------------------------
    btn = driver.find_element_by_xpath('//input[@id = "sdCd_F"]')
    driver.execute_script("arguments[0].click();", btn)
    #여자------------------------
#성별------------------------

#성------------------------
driver.find_element_by_css_selector("#eng2Fnm").send_keys(params.get('firstNm'))
#성------------------------

#이름------------------------
driver.find_element_by_css_selector("#eng1Fnm").send_keys(params.get('lastNm'))
#이름------------------------

#생년월일-년------------------------
select = Select(driver.find_element_by_css_selector('#btd_year'))
select.select_by_value(params.get('birthYyyy'))
#생년월일-년------------------------

#생년월일-월------------------------
select = Select(driver.find_element_by_css_selector('#btd_month'))
select.select_by_value(params.get('birthMm'))
#생년월일-월------------------------

#생년월일-일------------------------
select = Select(driver.find_element_by_css_selector('#btd_day'))
select.select_by_value(params.get('birthDd'))
#생년월일-일------------------------

#만료일-연------------------------
select = Select(driver.find_element_by_css_selector('#psExpr_year'))
select.select_by_value(params.get('passportExpiryYyyy'))
#만료일-연------------------------

#만료일-월------------------------
select = Select(driver.find_element_by_css_selector('#psExpr_month'))
select.select_by_value(params.get('passportExpiryMm'))
#만료일-월------------------------

#만료일-일------------------------
select = Select(driver.find_element_by_css_selector('#psExpr_day'))
select.select_by_value(params.get('passportExpiryDd'))
#만료일-일------------------------

#다음------------------------
driver.find_element_by_css_selector("#btnNext").click()
#다음------------------------

#STEP 04
#다른국가국민------------------------
if params.get('etcNationalityYn') == 'Y':
    #다른국가국민-네------------------------
    btn = driver.find_element_by_xpath('//input[@id = "diffPsNacdYn_Y"]')
    driver.execute_script("arguments[0].click();", btn)
    #다른국가국민-네------------------------

    #다른국가국민-국가------------------------
    select = Select(driver.find_element_by_css_selector('#multiNltDetCn'))
    select.select_by_value(params.get('anotherNationality'))
    #다른국가국민-국가------------------------
else:
    #다른국가국민-아니요------------------------
    btn = driver.find_element_by_xpath('//input[@id = "diffPsNacdYn_N"]')
    driver.execute_script("arguments[0].click();", btn)
    #다른국가국민-아니요------------------------
#다른국가국민------------------------

#휴대전화국가번호------------------------
select = Select(driver.find_element_by_css_selector('#countryCode'))
select.select_by_value(params.get('mobileNationCd'))
#휴대전화국가번호------------------------

#휴대전화번호------------------------
driver.find_element_by_css_selector("#hptel").send_keys(params.get('mobileNumber'))
#휴대전화번호------------------------

#과거대한민국방문------------------------
btn = driver.find_element_by_xpath('//input[@id = "pastKorVisYn_Y"]')
driver.execute_script("arguments[0].click();", btn)
#과거대한민국방문------------------------

#입국목적------------------------
select = Select(driver.find_element_by_css_selector('#entPurpCd'))
entPurpCd = params.get('enterReason')
select.select_by_value(entPurpCd)
if entPurpCd == "01":
    if 1 : #개별여행
        btn = driver.find_element_by_xpath('//input[@id = "grpTrYn_N"]')
        driver.execute_script("arguments[0].click();", btn)
    else: #단체여행
        btn = driver.find_element_by_xpath('//input[@id = "grpTrYn_Y"]')
        driver.execute_script("arguments[0].click();", btn)
elif entPurpCd == "99":
    driver.find_element_by_css_selector("#entPurpCn").send_keys("가나나나")
else:
    select.select_by_value(entPurpCd)
#입국목적------------------------

# driver.find_element_by_css_selector("#btnZip").click()
# time.sleep(2)
# #우편번호
# driver.find_element_by_css_selector("#keywordZipCode").send_keys("04524")
# driver.find_element_by_css_selector("#btnSerchZipCode").click()
# time.sleep(1)
# driver.find_element_by_xpath("//a[contains(@onclick,'addrSet')]").click()


# sojPrrplRnmBsAddr = driver.find_element_by_css_selector("#sojPrrplRnmBsAddr")

# #한국주소------------------------
# #우편번호------------------------
# driver.execute_script("document.getElementById('sojPrrplZip').setAttribute('type', 'text')");
# driver.find_element_by_css_selector("#sojPrrplZip").send_keys("04524")
# #우편번호------------------------
# #입국목적------------------------

# #영문주소------------------------
# driver.execute_script('document.getElementById("sojPrrplRnmBsAddr").removeAttribute("readonly")')
# driver.find_element_by_css_selector("#sojPrrplRnmBsAddr").send_keys("주소 주소")
# #영문주소------------------------

# #영문상세주소------------------------
# driver.find_element_by_css_selector("#sojPrrplRnmDetAddr").send_keys("101-101")
# #영문상세주소------------------------

# #한글주소------------------------
# driver.execute_script("document.getElementById('rnmBsAddr').setAttribute('type', 'text')");
# driver.find_element_by_css_selector("#rnmBsAddr").send_keys("서울특별시 중구 세종대로 110(태평로1가)")
# #한글주소------------------------
# #한국주소------------------------

#연락처------------------------
driver.find_element_by_css_selector("#sojPrrarTel").send_keys(params.get("enterTelephone"))
#연락처------------------------

#직업------------------------
select = Select(driver.find_element_by_css_selector('#etaOccpCd'))
etaOccpCd = params.get('job')
select.select_by_value(etaOccpCd)
if etaOccpCd == "99":
    driver.find_element_by_css_selector("#occpNm").send_keys(params.get('jobDescription'))
#직업------------------------

#질환------------------------
diseaseYn = params.get('diseaseYn')
if diseaseYn == "Y":
    btn = driver.find_element_by_xpath('//input[@id = "icdYn_Y"]')
    driver.execute_script("arguments[0].click();", btn)
    driver.find_element_by_css_selector("#icdRmkCn").send_keys(params.get('diseaseDescription'))
else:
    btn = driver.find_element_by_xpath('//input[@id = "icdYn_N"]')
    driver.execute_script("arguments[0].click();", btn)
#질환------------------------

#범죄------------------------
criminalYn = params.get('criminalYn')
if criminalYn == "Y":
    btn = driver.find_element_by_xpath('//input[@id = "crecYn_Y"]')
    driver.execute_script("arguments[0].click();", btn)
    driver.find_element_by_css_selector("#crecRmkCn").send_keys(params.get('criminalDescription'))
else:
    btn = driver.find_element_by_xpath('//input[@id = "crecYn_N"]')
    driver.execute_script("arguments[0].click();", btn)
#범죄------------------------

#안면사진파일추가------------------------
driver.find_element_by_css_selector("#btnAddFile").click()
#안면사진파일추가------------------------
time.sleep(3)

#안면사진파일열기------------------------
app = pywinauto.Application().connect(title="열기", class_name="#32770")
dlg = app.window(title_re="열기", class_name="#32770")
dlg.Edit1.SetEditText(IMG_FACE)
dlg.Button1.click()
#안면사진파일열기------------------------
time.sleep(3)

#확인------------------------
driver.find_element_by_css_selector("#btnPopConfirm").click()
#확인------------------------

#다음------------------------
# driver.find_element_by_link_text("다음").click()
# driver.find_element_by_css_selector("#btnPopConfirm").click()
#다음------------------------

#임시저장------------------------
driver.find_element_by_css_selector("#btnSaveTmp").click()
time.sleep(0.3)
driver.find_element_by_css_selector("#btnPopConfirm").click()
time.sleep(0.3)
driver.find_element_by_css_selector("#btnPopConfirm").click()
#임시저장------------------------


# #STEP 07
# driver.find_element_by_css_selector("#btnPay").click()

# print("")


messagebox.showinfo("작업완료", "입력내용을 확인하세요.\n\"확인\"버튼을 누르면 종료됩니다.")
print("")
