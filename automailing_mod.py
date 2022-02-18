from re import subn
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import pandas as pd

### 전송방식 선택 ###
mode = 0 # 0 : 개별 송부
         # 1 : 일괄 송부 

### 메일 제목
sub = "[한국로봇산업진흥원] 2022 로봇활용 제조혁신 지원사업 접수증 발급"

### 메일 내용 ###
txt = "안녕하십니까 한국로봇산업진흥원입니다.\n\n2022년 로봇활용 제조혁신 지원사업 신청해주시어, 접수증 발급드립니다.\n\n미비한 서류는 보완하시어 발송된 메일로 공고 마감일까지 회신하여 주시기 바랍니다.\n\n관련문의 : 한국로봇산업진흥원 제조로봇보급팀(053-210-9552~6)\n\n감사합니다."


df=pd.read_excel('C:/Users/TETRA/Desktop/접수증/접수현황(프로그램용).xlsx접수증리스트.xlsx')
mail = df['실무자 이메일']
file = df['접수증 경로']
corp = df['도입기업명']

maillist=mail.tolist()
maillist=str(','.join(maillist))
#print(maillist)

options= webdriver.ChromeOptions()
options.add_argument("--ignore-certificate-error")
options.add_argument("--ignore-ssl-errors")


driver = webdriver.Chrome('C:/Users/TETRA/Desktop/autocheck/chromedriver.exe',options=options)
driver.get("https://mail.kiria.org/index.ds")

#### 로그인 ########
id = driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/fieldset/div[1]/div[1]/input')
id.send_keys("factory")
pw=driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/fieldset/div[1]/div[2]/input')
pw.send_keys("wpwhgurtls1!")

driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/fieldset/div[2]/input').click()
#time.sleep(3)


    ######### 메일쓰기 클릭 ###########
element = driver.find_element_by_id("body_frame")
driver.switch_to.frame(element)
driver.find_element_by_xpath("/html/body/form[1]/div[3]/div/div[1]/div[1]/div/ul/li[1]/a/span").click()


if mode == 0:

    for i in range(len(df)):
        driver.switch_to.default_content()
        ###### 수신자, 제목 설정 ############
        element = driver.find_element_by_id("body_frame")
        driver.switch_to.frame(element)
        tomail = driver.find_element_by_css_selector('#mainDiv > div > table.mail_write > tbody > tr:nth-child(1) > td.rec > textarea')  ## 수신자 메일
        time.sleep(0.2)
        tomail.send_keys(mail[i]+',')
        subject = driver.find_element_by_css_selector('#mainDiv > div > table.mail_write > tbody > tr:nth-child(4) > td > textarea')
        time.sleep(0.2)
        subject.send_keys(sub+'_'+str(corp[i]))
        time.sleep(0.2)
        ## 제목


        #### 파일 첨부 ###
        
        element = driver.find_element_by_xpath("/html/body/form[1]/div[1]/div/div[3]/div/div[3]/div/div[3]/iframe")
        driver.switch_to.frame(element)
        driver.find_element_by_xpath('/html/body/form/div/table/tbody/tr[1]/td/input[3]').send_keys(str(file[i])) ## 첨부할 파일 경로
        


        ###### 메일 작성#######
        driver.switch_to.default_content()
        element = driver.find_element_by_id("body_frame")
        driver.switch_to.frame(element)

        element = driver.find_element_by_id("freeRTE")
        driver.switch_to.frame(element)
        driver.find_element_by_css_selector('body').send_keys(txt) # 내용 기입
        driver.switch_to.default_content() #처음 상태로 되돌아옴
        

        #### 보내기 ###

        element = driver.find_element_by_id("body_frame")
        driver.switch_to.frame(element)
        driver.find_element_by_xpath('/html/body/form[1]/div[1]/div/div[3]/div/div[2]/div/a[1]').click()
        time.sleep(1)


        driver.back()
        time.sleep(0.5)
        driver.back()
        time.sleep(0.5)
        #driver.back()
        ### 홈화면으로 ###
        #element = driver.find_element_by_name("topframe")
        #driver.switch_to.frame(element)
        #driver.find_element_by_xpath('/html/body/div/div[2]/a[1]').click()
        #driver.switch_to.default_content()
    # time.sleep(2)
        #driver.execute_script('javascript: movePage(\'home.ds\',\'home\');')

if mode == 1:
    driver.switch_to.default_content()
    ###### 수신자, 제목 설정 ############
    element = driver.find_element_by_id("body_frame")
    driver.switch_to.frame(element)
    driver.find_element_by_css_selector('#mainDiv > div > table.mail_write > tbody > tr:nth-child(1) > td.rec > textarea').send_keys(maillist)  ## 수신자 메일
    driver.find_element_by_css_selector('#mainDiv > div > table.mail_write > tbody > tr:nth-child(4) > td > textarea').send_keys(sub)
    ## 제목


    #### 파일 첨부 ###
    """    
    element = driver.find_element_by_xpath("/html/body/form[1]/div[1]/div/div[3]/div/div[3]/div/div[3]/iframe")
    driver.switch_to.frame(element)
    driver.find_element_by_xpath('/html/body/form/div/table/tbody/tr[1]/td/input[3]').send_keys(str(file[i])) # 첨부파일 경로
        """


    ###### 메일 작성#######
    driver.switch_to.default_content()
    element = driver.find_element_by_id("body_frame")
    driver.switch_to.frame(element)

    element = driver.find_element_by_id("freeRTE")
    driver.switch_to.frame(element)
    driver.find_element_by_css_selector('body').send_keys(txt)
    driver.switch_to.default_content() #처음 상태로 되돌아옴
        

    #### 보내기 ###

    element = driver.find_element_by_id("body_frame")
    driver.switch_to.frame(element)
    driver.find_element_by_xpath('/html/body/form[1]/div[1]/div/div[3]/div/div[2]/div/a[1]').click()
    time.sleep(1)
