from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException

import re, os
import time
import platform
import random


def login(id, name, university):
    driver.find_element_by_class_name("chosen-single").click()
    search_box = driver.find_element_by_xpath("//div[@class='chosen-search']/input")
    search_box.click()
    search_box.send_keys(university)
    driver.find_element_by_xpath("//ul[@class='chosen-results']/li").click()

    id_box = driver.find_element_by_id("UniqueKey1")
    name_box = driver.find_element_by_name("UserName")

    id_box.send_keys(id)
    name_box.send_keys(name)
    name_box.send_keys(Keys.RETURN)

    WebDriverWait(driver, 10)
    
def downloadPDF(id, name, cerValue):
    import requests
    import json

    url = "https://safetyedu.org/Account/AjxAgreementChk"

    data = {"AgencyNo": 18, "LoginType": 1, "UniqueKey": id, "UserName": name}

    # res = requests.post(url, data)
    session = requests.session()
    session.post(url, data=data)

    res = session.get("http://safetyedu.org/Edu/Certificate")

    data = {"cerValue" : cerValue}
    pdf_data = session.post("https://safetyedu.org/Edu/CompletPdf", data=data)
    pdf_data = json.loads(pdf_data.content)

    pdf = session.get("https://safetyedu.org/Edu/FileDownLoad?ScheduleName=" + pdf_data["ScheduleName"] + " &PdfName=" + pdf_data["PdfName"])

    with open("./" + name + id + ".pdf", 'wb') as f:
        f.write(pdf.content)
    print("\n" + os.getcwd() + "\\" + name + id + ".pdf " + "저장되었습니다.\n")

university = input("University : ")
studentID = input("Student ID : ")
studentName = input("Name : ")

if platform.system() == "Windows":
    os.system("cls")
else:
    os.system("clear")


file_list = os.listdir(os.getcwd())
driver_path = None

for i in file_list:
    if(re.search("chromedriver", i)):
        driver_path = "./chromedriver"

if not driver_path:
    print("chromedriver 파일을 찾을 수 없습니다. 직접 경로를 입력해주세요.")
    driver_path = input("Google Chrome Driver Path : ")

driver = webdriver.Chrome(driver_path.replace("\\", "/"))

main_window = driver.window_handles[0]

driver.get("https://safetyedu.org/Account/LogOn")
WebDriverWait(driver, 10)
login(studentID, studentName, university)

driver.get("https://safetyedu.org/Edu/OnLineEdu")
listen_flag = False

try:
    driver.find_element_by_xpath("//input[@class='btn_ExAlramPop1']").click()
except NoSuchElementException:
    listen_flag = True

if listen_flag:
    isListen = driver.find_element_by_xpath("//p[@class='edu_FireStatus']/label").text

    if(isListen == "과목선택"):
        driver.find_element_by_id("btnMappingContent").click()
        time.sleep(2)
        subject_tables = driver.find_elements_by_xpath("//table[@class='col_table scroll_table fht-table fht-table-init']/tbody/tr")

        for tr in subject_tables:
            select_status = driver.find_element_by_class_name("selectOSpan").text.split(" / ")
            if (int(select_status[0]) < int(select_status[1])):
                checkbox = tr.find_element_by_xpath(".//td[@class='tac']")
                checkbox.click()

        # 과목 선택 완료
        driver.find_element_by_class_name("setBtn").click()

    elif (isListen == "교육수강"):
        for i in range(6):
            try:
                # HTML 페이지에서 수강하기 버튼을 찾음
                driver.find_element_by_xpath("//table[@class='edufireTable']/tbody/tr/td/input[@value='수강하기']").click()

                # 새로운 창으로 기다림
                while(len(driver.window_handles) < 2):
                    pass
                
                # 새창으로 창을 전환함
                driver.switch_to.window(driver.window_handles[1])
                driver.set_window_size(driver.get_window_size()["width"], driver.get_window_size()["height"] + 100)
                time.sleep(3)

                # 해당 동영상의 정보를 얻음
                # nowPage : 현재 진행중인 페이지, endPage : 마지막 페이지
                # nowTime : 내가 듣고 있는 위치, endTime : 동영상의 끝시간
                nowPage = int(driver.find_element_by_xpath("//div[@class='pageNum cPageNum']").text)
                endPage = int(driver.find_element_by_xpath("//div[@class='pageNum tPageNum']").text)
                nowTime = driver.find_element_by_xpath("//div[@class='cTime']").text
                endTime = driver.find_element_by_xpath("//div[@class='dTime']").text

                # 한개의 강의를 다 들었는지 확인
                while(nowPage != endPage or nowTime != endTime):
                    if platform.system() == "Windows":
                        os.system("cls")
                    else:
                        os.system("clear")
                    print("=" * 40)
                    print("강의 페이지 : " + str(nowPage) + " / " + str(endPage))
                    # 시간을 매번 갱신해줌
                    nowTime = driver.find_element_by_xpath("//div[@class='cTime']").text
                    endTime = driver.find_element_by_xpath("//div[@class='dTime']").text
                    print("강의 시간 : " + nowTime + " / " + endTime)
                    print("=" * 40)

                    try:
                        # 중간에 퀴즈풀기 등과 같은 동영상을 건너뛰기위해서 재생버튼을 수시로 눌러줌
                        driver.find_element_by_xpath("//div[@class='ctrlBtn playBtn on']").click()
                    except:
                        pass
                
                    # 강의가 끝났다면
                    if(nowTime == endTime and nowPage != endPage):
                        print("%%%% 강의가 끝났습니다. 다음으로 넘어갑니다. %%%%")
                        driver.find_element_by_xpath("//div[@class='moveBtn nextPageBtn']").click()
                        # 간혹 사라지는 하단의 동영상 컨트롤바를 위해 창 크기를 한번 조정해준다.
                        if (driver.get_window_size()["height"] % 2 == 0):
                            driver.set_window_size(driver.get_window_size()["width"], driver.get_window_size()["height"] + 1)
                        else:
                            driver.set_window_size(driver.get_window_size()["width"], driver.get_window_size()["height"] - 1)
                        time.sleep(1)
                        # 다음 동영상의 정보를 미리 파악한다.
                        nowPage = int(driver.find_element_by_xpath("//div[@class='pageNum cPageNum']").text)
                        nowTime = driver.find_element_by_xpath("//div[@class='cTime']").text
                        endTime = driver.find_element_by_xpath("//div[@class='dTime']").text
                    time.sleep(1)
                driver.close()
                driver.switch_to.window(main_window)
            except NoSuchElementException:
                print("수강하기가 없습니다.")

    time.sleep(1)
    try:
        driver.find_element_by_xpath("//input[@class='btn_ExAlramPop1']").click()
    except NoSuchElementException:
        pass

time.sleep(1)
try:
    myPoint = 0
    while (myPoint < 60):
        exam = driver.find_elements_by_xpath("//table[@id='Exam_tblList']/tbody/tr")

        title_list = [tr for idx, tr in enumerate(exam) if idx % 2 == 0]
        answer_list = [tr for idx, tr in enumerate(exam) if idx % 2 == 1]

        rightAnswer = {}

        for title, answer in zip(title_list, answer_list):
            title = title.find_element_by_xpath(".//td[@class='tal']").text
            answer_chose = answer.find_elements_by_xpath(".//input[@class='radioAnswer']")
            if title in rightAnswer:
                ansTitle = answer.find_elements_by_xpath(".//td[@class='tal']/label")
                for idx, ans in enumerate(ansTitle):
                    if ans == rightAnswer[title]:
                        answer_chose[idx].click()
            else:
                random.choice(answer_chose).click()

        driver.find_element_by_xpath("//input[@title='제출하기']").click()

        time.sleep(0.5)
        alert = Alert(driver)
        myPoint = int(re.findall("[0-9]+", alert.text)[-1])
        alert.accept()

        endExam = driver.find_elements_by_xpath("//table[@id='Exam_tblList']/tbody/tr")
        title_list = [tr for idx, tr in enumerate(endExam) if idx % 2 == 0]
        answer_list = [tr for idx, tr in enumerate(exam) if idx % 2 == 1]

        for title, answer in zip(title_list, answer_list):
            atitle = title.find_element_by_xpath(".//td[@class='tal']").text
            ansNumber = int(title.find_element_by_xpath(".//span[@class='CorrectAnswer']").text[0])
            ansTitle = answer.find_elements_by_xpath(".//td[@class='tal']/label")[ansNumber - 1].text
            rightAnswer[atitle] = ansTitle

        driver.find_element_by_xpath("//input[@title='닫기']").click()
        time.sleep(1)
        try:
            driver.find_element_by_xpath("//input[@class='btn_ExAlramPop1']").click()
        except NoSuchElementException:
            pass
        time.sleep(2)
except:
    print("이미 평가를 통과하셨거나 다른 문제가 생겼습니다.") 

driver.get("https://safetyedu.org/Edu/Certificate")
time.sleep(1)
checkList = driver.find_elements_by_xpath("//table[@class='eduLabTable']/tbody/tr")
cerValue = checkList[1].find_element_by_xpath(".//input[@name='chCertificate']").get_attribute("value")

print(cerValue)
print(list(cerValue))
input()
downloadPDF(studentID, studentName, cerValue)

print("안전교육을 모두 이수하셨습니다.")
print("5초 후에 자동으로 종료됩니다.")
time.sleep(5)
driver.quit()