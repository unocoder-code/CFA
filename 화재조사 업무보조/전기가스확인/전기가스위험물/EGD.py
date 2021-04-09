from json import load, loads
from re import search
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import *
import time
from selenium.webdriver.remote.utils import format_json
from tkinter import *

def log_in():
    #로그인
    driver.find_element_by_id("_easyui_textbox_input4").click()
    driver.find_element_by_id("_easyui_combobox_i1_12").click()
    driver.find_element_by_id('_easyui_textbox_input1').send_keys('0105108374311')
    driver.find_element_by_id('_easyui_textbox_input2').send_keys('0000')
    time.sleep(0.5)
    driver.find_element_by_id("btnLogin").click()
    time.sleep(1)
    check_loading()

def click(path):
    driver.find_element_by_xpath(path).click()

def sendkeyid(path,key):
    driver.find_element_by_id(path).send_keys(key)

def sendkey(path,key):
    driver.find_element_by_xpath(path).send_keys(key)


def check_loading():
    try:
        html_raw = driver.page_source
        html = BeautifulSoup(html_raw, 'html.parser')
        loadingcode = html.find(id="fsseProgress")
        try:
            load_attrs = loadingcode.attrs
            if 'in' in load_attrs['class']:
                print ('*', end='')
                time.sleep(1)
                check_loading()
            else:
                time.sleep(1)
                print('\nEnd Check loading')
        except:
            print("Error")
            time.sleep(100)
            pass
    except:
        check_loading()
        pass

def page_loading():
    try:
        html_raw = driver.page_source
        html = BeautifulSoup(html_raw, 'html.parser')
        page_code = html.find(class_="datagrid-mask")
        if page_code == None:
            time.sleep(1)
            print('\nEnd Page loading')
        else:
            print ('*', end='')
            time.sleep(1)
            page_loading()
    except:
        time.sleep(100)
        pass

def 저장_확인_클릭_로딩():
    while True:
        try:
            click('//*[@id="b_Insert"]')
            break
        except:
            time.sleep(1)
            pass
    time.sleep(1)
    while True:
        try:
            click('//*[@id="btn_ok_c"]')
            break
        except:
            time.sleep(1)
            pass
    time.sleep(1)
    while True:
        try:
            click('//*[@id="btn_ok_a"]')
            break
        except:
            time.sleep(1)
            pass
    time.sleep(1)
    check_loading()

def 전기_가스_위험물():
    for i in range(5,8):
        click('//*[@id="tab_title{}"]'.format(i))
        저장_확인_클릭_로딩()

def pagination(): #25단위로 확장
    click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[1]/select/option[4]')
    page_loading()

def choose_condition(): 
    #충주시 & 2020년도 & 조사반C 선택
    click("//option[@value='4313000000']") 
    time.sleep(0.1)
    click('//*[@id="srch_scxm"]/option[5]')
    time.sleep(0.1)
    click('//*[@id="srch_prgss_stt"]/option[4]')
    time.sleep(0.1)
    driver.find_element_by_class_name("btn_search").click()

def init_page(page):
    real = page - 1
    x = int( real / 10)
    y = int((real%10) / 4)
    z = (real%10) % 4
    for a in range (x):
        click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[8]/a/span')
        page_loading()
    for a in range (y):
        click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[6]/a[{}]'.format(10))
        page_loading()
    click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[6]/a[{}]'.format(6+z))
    page_loading()

def 사이트_로그인_실행():
    global driver
    driver = webdriver.Chrome(chromedriver)
    driver.get('http://10.175.105.14/FSSE/lgn/login2.do')
    log_in()
    check_loading()
    pagination()
    choose_condition()
    check_loading()
    click('//*[@id="datagrid-row-r2-2-0"]/td[2]')
    check_loading()
    while True:
        try:
            click('//*[@id="limitCont"]/div[5]/a[1]')
            break
        except:
            time.sleep(1)
            pass
    check_loading()

def 한리스트작동():
    c = 7
    for i in range(25):
        print(c,'-',i)
        try:
            click('//*[@id="datagrid-row-r1-2-{}"]'.format(i))
        except:
            click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[11]/a/span/span[2]')
            page_loading()
            click('//*[@id="datagrid-row-r1-2-{}"]'.format(i))
        check_loading()
        전기_가스_위험물()
        click('//*[@id="limitCont"]/div[5]/a[1]') #목록
        check_loading()

def 페이지8_10작동():
    b = 7
    click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[6]/a[8]')
    page_loading()
    한리스트작동()
    
    for i in range (2):
        click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[6]/a[7]')
        page_loading()
        한리스트작동()
        b+=1


def 페이지연속작동():
    while True:
        click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[6]/a[7]')
        page_loading()
        한리스트작동()

#다음페이지(10단위 점프)
#click('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[8]/a/span/span[2]')

def gui_program():
    root = Tk()
    root.title("1전기가스프로그램1")
    root.geometry("400x400")
    사이트_로그인 = Button(root, width=20,height=3, text="사이트_로그인", command=사이트_로그인_실행)
    사이트_로그인.pack()
    페이지1_10작동버튼 = Button(root, width=20,height=3, text="페이지8_10작동", command=페이지8_10작동)
    페이지1_10작동버튼.pack()
    root.mainloop()

if __name__ == '__main__':
    chromedriver = 'C:/Users/win10/Desktop/자동화 프로그램/전기가스위험물/chromedriver.exe'
    gui_program()