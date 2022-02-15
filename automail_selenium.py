import py_compile
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import pandas as pd


df=pd.read_excel('D:/접수증/접수증리스트.xlsx')
mail = df['총괄이메일']
file = df['접수증 경로']

options= webdriver.ChromeOptions()
options.add_argument("--ignore-certificate-error")
options.add_argument("--ignore-ssl-errors")


driver = webdriver.Chrome('D:/automail/chromedriver.exe',options=options)
driver.get("https://mail.kiria.org/index.ds")

#### 로그인 ########
id = driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/fieldset/div[1]/div[1]/input')
id.send_keys("sjyet96")
pw=driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/fieldset/div[1]/div[2]/input')
pw.send_keys("Songjiyou!1")

driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/form/fieldset/div[2]/input').click()
#time.sleep(3)


    ######### 메일쓰기 클릭 ###########
element = driver.find_element_by_id("body_frame")
driver.switch_to.frame(element)
driver.find_element_by_xpath("/html/body/form[1]/div[3]/div/div[1]/div[1]/div/ul/li[1]/a/span").click()

for i in range(len(df)):
    driver.switch_to.default_content()
    ###### 수신자, 제목 설정 ############
    element = driver.find_element_by_id("body_frame")
    driver.switch_to.frame(element)
    driver.find_element_by_css_selector('#mainDiv > div > table.mail_write > tbody > tr:nth-child(1) > td.rec > textarea').send_keys(mail[i]+',')  ## 수신자 메일
    driver.find_element_by_css_selector('#mainDiv > div > table.mail_write > tbody > tr:nth-child(4) > td > textarea').send_keys("[한국로봇산업진흥원] 2022년 로봇활용 제조혁신 지원사업 접수증 송부")
    ## 제목


    #### 파일 첨부 ###
    element = driver.find_element_by_xpath("/html/body/form[1]/div[1]/div/div[3]/div/div[3]/div/div[3]/iframe")
    driver.switch_to.frame(element)
    driver.find_element_by_xpath('/html/body/form/div/table/tbody/tr[1]/td/input[3]').send_keys(str(file[i]))
    driver.switch_to.default_content()


    ###### 메일 작성#######

    element = driver.find_element_by_id("body_frame")
    driver.switch_to.frame(element)

    element = driver.find_element_by_id("freeRTE")
    driver.switch_to.frame(element)
    driver.find_element_by_css_selector('body').send_keys("내용입력\n내용 두번째 줄 \n\n두줄띄우기")
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
    