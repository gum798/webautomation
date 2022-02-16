import json
import os
import shutil
import sys
import time
import traceback
from ftplib import FTP
from tkinter import *
from tkinter import messagebox
from urllib import parse

import pywinauto
import win32api
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import chromedriver_autoinstaller

chromedriver_autoinstaller.install(path="C:/keta/")


bDebug = False

# if bDebug:
#     print("")
#     # messagebox.showinfo("파라미터", "디버그모드") #, parent=topWindow)
# else:
#     # messagebox.showinfo("파라미터", "운영모드") #, parent=topWindow)
#     if sys.argv[-1] != 'asadmin':
#         script = os.path.abspath(sys.argv[0])
#         params = ' '.join(sys.argv[1:] + ['asadmin'])
#         # win32api.ShellExecute(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
#         # messagebox.showinfo(script, params) #, parent=topWindow)
#         # win32api.ShellExecute(0, 'runas', script, params,'c:\keta',0)
#         win32api.ShellExecute(0, None, script, params, None, 0)
#         sys.exit()


root = Tk()
root.wm_attributes("-topmost", 1)
# topWindow = Toplevel(root)
root.withdraw()

if len(sys.argv)>1:
    # messagebox.showinfo("파라미터", sys.argv[1]) #, parent=topWindow)
    print("")
else:
    messagebox.showerror("ERROR", "파라미터가 없습니다.")
    sys.exit()

try:
    if sys.argv[1][-1]=="/":
        params = json.loads(parse.unquote(sys.argv[1].replace("keta://","")[:-1]))
    else:
        params = json.loads(parse.unquote(sys.argv[1].replace("keta://","")))
except:
    messagebox.showerror("ERROR", f"파라미터 형식에 문제가 있습니다.\n{traceback.format_exc()}")
    sys.exit()

try:
    def getImagePath(uploadFileList, fileType):
        for file in uploadFileList:
            if file.get("fileType").upper() == fileType.upper():
                return file.get("uploadPath")
    imgPassport = getImagePath(params.get("uploadFileList"), "IMG_PASSPORT")
    imgFace = getImagePath(params.get("uploadFileList"), "IMG_FACE")
except:
    messagebox.showerror("ERROR", f"이미지 정보를 찾을 수 없습니다.\n{traceback.format_exc()}")
    sys.exit()

if imgPassport == None or imgFace == None:
    messagebox.showerror("ERROR", "이미지 정보를 찾을 수 없습니다.")
    sys.exit()

try:
    shutil.rmtree(r"C:\\keta\\tmp")
except:
    # messagebox.showwarning("WARNING", "C:\\keta\\tmp 폴더를 삭제할 수 없습니다. 수동으로 삭제해 주세요")
    print(sys.exc_info())
    print("C:\\keta\\tmp 폴더를 삭제할 수 없습니다. 수동으로 삭제해 주세요")

try:
    def download_file(ftp, dir, file):
        ftp.cwd("/")
        ftp.cwd(dir[0])
        ftp.cwd(dir[1])
        # download_path = os.path.abspath("./tmp/")
        download_path = "C:/keta/tmp/"
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



    #FTP접속
    if bDebug:
        ftp = FTP('121.78.118.213')
        # ftp = FTP('121.78.145.131')
    else:
        ftp = FTP('121.78.145.131')

    ftp.login("ftpketa","ftpketa!234")
    print(ftp.retrlines('LIST'))
    srcDir, srcFileName = path_parser(imgPassport)
    IMG_PASSPORT = download_file(ftp, srcDir, srcFileName)
    srcDir, srcFileName = path_parser(imgFace)
    IMG_FACE = download_file(ftp, srcDir, srcFileName)

    ftp.quit()
except:
    messagebox.showerror("ERROR", f"FTP 통신 에러\n{traceback.format_exc()}")
    ftp.quit()
    sys.exit()

try:
    #크롬드라이버 로딩
    # driver= webdriver.Chrome()
    
    # options = webdriver.ChromeOptions()
    # options.add_experimental_option("detach", True)

    if getattr(sys, 'frozen', False):
        chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
        # driver = webdriver.Chrome(chrome_options=options, executable_path=chromedriver_path)
        driver = webdriver.Chrome(executable_path=chromedriver_path)
    else:
        # driver = webdriver.Chrome(chrome_options=options)
        driver = webdriver.Chrome()
        # driver = webdriver.Chrome(driver_path)
        

    url = 'https://www.k-eta.go.kr/portal/apply/viewstep1.do'
    driver.get(url)
    driver.maximize_window()
    action = ActionChains(driver)
except:
    messagebox.showerror("ERROR", f"크롬 드라이버 로딩 에러\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #STEP 01
    #대륙선택------------------------
    # driver.find_element_by_css_selector("#etaCnttGbcd").click()
    # action.send_keys("AMERICA").perform()
    select = Select(driver.find_element_by_css_selector('#etaCnttGbcd'))
    select.select_by_value(params.get('etaCnttGbcd'))
    action.reset_actions()
    driver.find_element_by_css_selector("#natCd").click()
    #대륙선택------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\n대륙정보 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #국적선택------------------------
    select = Select(driver.find_element_by_css_selector('#natCd'))
    select.select_by_value(params.get('passportPublishCountry'))
    #국적선택------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\n국적정보 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()


#신청불가------------------------
exit_flag = False
try:
#이미작성된내용-취소------------------------
    ok = driver.find_element_by_css_selector(".popBtn2")
    if ok != None:
        messagebox.showinfo("신청불가", f"\"확인\"버튼을 누르면 종료됩니다.\n{traceback.format_exc()}")
        exit_flag = True
except:
    print("")
if exit_flag:
    driver.quit()
    sys.exit()
#신청불가------------------------

try:
    #필수개인정보수집동의------------------------
    agree1 = driver.find_element_by_css_selector("#dAgree1")
    action.move_to_element(agree1).click().perform()
    action.reset_actions()
    #필수개인정보수집동의------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\n필수개인정보수집동의 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #선택개인정보수집동의------------------------
    agree2 = driver.find_element_by_css_selector("#dAgree2")
    action.move_to_element(agree2).click().perform()
    action.reset_actions()
    #선택개인정보수집동의------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\n선택개인정보수집동의 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #민감정보처리동의------------------------
    agree3 = driver.find_element_by_css_selector("#dAgree3")
    action.move_to_element(agree3).click().perform()
    action.reset_actions()
    #민감정보처리동의------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\n민감정보처리동의 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #고유식별정보처리동의------------------------
    agree4 = driver.find_element_by_css_selector("#dAgree4")
    action.move_to_element(agree4).click().perform()
    action.reset_actions()
    #고유식별정보처리동의------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\고유식별정보처리동의 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #약관동의------------------------
    agree5 = driver.find_element_by_css_selector("#dAgree5")
    action.move_to_element(agree5).click().perform()
    action.reset_actions()
    #약관동의------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\약관동의 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #다음------------------------
    driver.find_element_by_css_selector(".btnBasic2.popupCheck").click()
    #다음------------------------
except:
    messagebox.showerror("ERROR", f"STEP.01 진행중 오류 발생\다음버튼 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

time.sleep(1)

#STEP 02
try:
    #여권번호------------------------
    driver.find_element_by_css_selector("#psno").send_keys(params.get('passportNo'))
    #여권번호------------------------
except:
    messagebox.showerror("ERROR", f"STEP.02 진행중 오류 발생\n여권번호 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #이메일주소------------------------
    driver.find_element_by_css_selector("#emlAddr").send_keys(params.get('email'))
    #이메일주소------------------------
except:
    messagebox.showerror("ERROR", f"STEP.02 진행중 오류 발생\n이메일주소 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #확인------------------------
    driver.find_element_by_css_selector(".btnInput1").click()
    #확인------------------------
except:
    messagebox.showerror("ERROR", f"STEP.02 진행중 오류 발생\n확인 버튼 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

time.sleep(1)

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

def connect_dialog(title, classname, path):
    for i in range(30):
        try:
            app = pywinauto.Application().connect(title=title, class_name=classname)
        except:
            break
        dlg = app.window(title_re=title, class_name=classname)
        dlg.Edit1.SetEditText(path)
        time.sleep(1)
        dlg.Button1.click()
        time.sleep(0.5)

try:
    # 여권사진 업로드
    # driver.find_element_by_css_selector(".irx_filetree_list.irx_listGrid").send_keys("C:\\Users\\jhseo\\Desktop\\여권.png")
    #파일추가------------------------
    driver.find_element_by_css_selector("#btnAddFile").click()
    time.sleep(3)
    #파일추가------------------------

    #파일열기------------------------
    connect_dialog("열기", "#32770", IMG_PASSPORT)
    #파일열기------------------------
    time.sleep(3)
except:
    messagebox.showerror("ERROR", f"여권사진 업로드 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
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
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n성별 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #성------------------------
    driver.find_element_by_css_selector("#eng2Fnm").send_keys(params.get('firstNm'))
    #성------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n성 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #이름------------------------
    driver.find_element_by_css_selector("#eng1Fnm").send_keys(params.get('lastNm'))
    #이름------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n이름 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #생년월일-년------------------------
    select = Select(driver.find_element_by_css_selector('#btd_year'))
    select.select_by_value(params.get('birthYyyy'))
    #생년월일-년------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n생년월일-년 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #생년월일-월------------------------
    select = Select(driver.find_element_by_css_selector('#btd_month'))
    select.select_by_value(params.get('birthMm'))
    #생년월일-월------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n생년월일-월 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #생년월일-일------------------------
    select = Select(driver.find_element_by_css_selector('#btd_day'))
    select.select_by_value(params.get('birthDd'))
    #생년월일-일------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n생년월일-일 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #만료일-연------------------------
    select = Select(driver.find_element_by_css_selector('#psExpr_year'))
    select.select_by_value(params.get('passportExpiryYyyy'))
    #만료일-연------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n만료일-연 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #만료일-월------------------------
    select = Select(driver.find_element_by_css_selector('#psExpr_month'))
    select.select_by_value(params.get('passportExpiryMm'))
    #만료일-월------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n만료일-월 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #만료일-일------------------------
    select = Select(driver.find_element_by_css_selector('#psExpr_day'))
    select.select_by_value(params.get('passportExpiryDd'))
    #만료일-일------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n만료일-일 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #다음------------------------
    driver.find_element_by_css_selector("#btnNext").click()
    #다음------------------------
except:
    messagebox.showerror("ERROR", f"STEP.03 진행중 오류 발생\n다음 버튼 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

time.sleep(1)
#STEP 04

try:
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
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n다른국가국민 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #휴대전화국가번호------------------------
    select = Select(driver.find_element_by_css_selector('#countryCode'))
    select.select_by_value(params.get('mobileNationCd'))
    #휴대전화국가번호------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n휴대전화국가번호 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #휴대전화번호------------------------
    driver.find_element_by_css_selector("#hptel").send_keys(params.get('mobileNumber'))
    #휴대전화번호------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n휴대전화번호 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #과거대한민국방문------------------------
    btn = driver.find_element_by_xpath('//input[@id = "pastKorVisYn_Y"]')
    driver.execute_script("arguments[0].click();", btn)
    #과거대한민국방문------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n과거대한민국방문 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #입국목적------------------------
    select = Select(driver.find_element_by_css_selector('#entPurpCd'))
    entPurpCd = params.get('enterReason')
    # select.select_by_value(entPurpCd)
    if entPurpCd == "01":
        #개별여행
        select.select_by_value(entPurpCd)
        btn = driver.find_element_by_xpath('//input[@id = "grpTrYn_N"]')
        driver.execute_script("arguments[0].click();", btn)
    elif entPurpCd == "98":
        #단체여행
        select.select_by_value("01")
        btn = driver.find_element_by_xpath('//input[@id = "grpTrYn_Y"]')
        driver.execute_script("arguments[0].click();", btn)
    elif entPurpCd == "99":
        select.select_by_value(entPurpCd)
        driver.find_element_by_css_selector("#entPurpCn").send_keys(params.get("enterReasonDetail"))
    else:
        select.select_by_value(entPurpCd)
    #입국목적------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n입국목적 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()


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

try:
    #연락처------------------------
    driver.find_element_by_css_selector("#sojPrrarTel").send_keys(params.get("enterTelephone"))
    #연락처------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n연락처 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #직업------------------------
    select = Select(driver.find_element_by_css_selector('#etaOccpCd'))
    etaOccpCd = params.get('job')
    select.select_by_value(etaOccpCd)
    if etaOccpCd == "99":
        driver.find_element_by_css_selector("#occpNm").send_keys(params.get('jobDescription'))
    #직업------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n직업 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
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
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n질환 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
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
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n범죄 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
    #안면사진파일추가------------------------
    driver.find_element_by_css_selector("#btnAddFile").click()
    #안면사진파일추가------------------------
    time.sleep(3)

    #안면사진파일열기------------------------
    connect_dialog("열기", "#32770", IMG_FACE)
    #안면사진파일열기------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n안면사진파일 업로드 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

time.sleep(3)

try:
    #확인------------------------
    driver.find_element_by_css_selector("#btnPopConfirm").click()
    #확인------------------------
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n확인 버튼 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

try:
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
except:
    messagebox.showerror("ERROR", f"STEP.04 진행중 오류 발생\n임시저장 오류\n{traceback.format_exc()}")
    driver.quit()
    sys.exit()

# #STEP 07
# driver.find_element_by_css_selector("#btnPay").click()

# print("")

messagebox.showinfo("작업완료", "매크로 동작이 완료되었습니다.\n*입력 내용 확인 및 추가 작업 진행 후 \"확인\" 버튼을 눌러주세요\n\"확인\"버튼을 누르면 브라우저가 닫힙니다.")
print("")

driver.quit()
print("")

