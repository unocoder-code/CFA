from json import load, loads
from re import search
from selenium import webdriver
from bs4 import BeautifulSoup
import openpyxl
import time

from soupsieve.css_parser import FLG_HTML, VALUE

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
def check_file():
    html_raw = driver.page_source
    html = BeautifulSoup(html_raw, 'html.parser')
    #for i in range(1,5):
    for i in range(4,0,-1):
        file_code = html.find(id="tr_file-{}".format(i))
        file_attrs = file_code.attrs
        print(file_attrs)
        if 'class' in file_attrs:
            print('inserted')
        else:
            print('Not image')
            Append_Not_image_list(html)
            break
def Append_Not_image_list(sorce): #사진 없을 경우 엑셀에 목록 추가
    #건물 이름 : building_name
    building_name_code = sorce.find(id="tab1_fon")
    print(building_name_code)
    building_name_attrs = building_name_code.attrs
    building_name = building_name_attrs['value']
    
    #건물 주소 : building_address
    building_address_code = sorce.find(id="searchAdress")
    print(building_address_code)
    building_address_attrs = building_address_code.attrs
    building_address = building_address_attrs['value']
    save_excel(building_name,building_address)

def save_excel(name,address):
    wb = openpyxl.load_workbook('화재안정정보조사 시스템 이미지 없.xlsx')
    sheet1 = wb.active
    sheet1.append([name,address])
    wb.save('화재안정정보조사 시스템 이미지 없.xlsx')
def click_mouse_id(id):
    driver.find_element_by_id(id).click()
def send_key_id(id,key):
    driver.find_element_by_id(id).send_keys(key)
def click_mouse_class(name):
    driver.find_element_by_class_name(name).click()
def click_mouse_xpath(path):
    driver.find_element_by_xpath(path).click() 
def log_in():
    #로그인
    click_mouse_id("_easyui_textbox_input4")
    click_mouse_id("_easyui_combobox_i1_12")
    send_key_id('_easyui_textbox_input1','')
    send_key_id('_easyui_textbox_input2','')
    time.sleep(0.5)
    click_mouse_id("btnLogin")
    time.sleep(1)
def choose_condition(): 
    #충주시 & 2020년도 & 조사반C 선택
    click_mouse_id("addr2")
    click_mouse_xpath("//option[@value='4313000000']") 
    time.sleep(0.1)

    click_mouse_id("srch_year_select")
    click_mouse_xpath("//option[@value='2020']")

    click_mouse_id("BTAR_GRAD_CD")
    click_mouse_xpath("//option[@value='C']")
    click_mouse_class("btn_search")
def pagination(): #25단위로 확장
    click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[1]/select/option[4]')
    page_loading()
def next_page(): #리스트 다음페이지로
    click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[8]/a/span')
    page_loading()
def click_index(): #목록 클릭
    click_mouse_xpath('//*[@id="limitCont"]/div[5]/a[1]')
    print("go index")
    check_loading()
    print("fin")
def enter_list_first():
    click_mouse_xpath('//*[@id="datagrid-row-r2-2-0"]')
    time.sleep(0.1)
    check_file()
    print("fin check file")
    time.sleep(0.1)
    try:
        click_index()
    except:
        click_index()
        pass
    time.sleep(0.1)
def enter_list(num):
    click_mouse_xpath('//*[@id="datagrid-row-r1-2-{}"]'.format(num))
    time.sleep(0.1)
    check_file()
    time.sleep(0.1)
    click_index()
    time.sleep(0.1)
def page_select(num):
    click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[6]/a[{}]/span/span'.format(num))
    page_loading()
def init_page(page):
    real = page - 1
    x = int( real / 10)
    y = int((real%10) / 4)
    z = (real%10) % 4
    for a in range (x):
        next_page()
    for a in range (y):
        page_select(10)
    page_select(6+z)
def check_page():
    for i in range (25):
        try:
            enter_list(i)
        except:
            save_excel('Error!!',0)
            time.sleep(100)
            pass
if __name__ == '__main__':
    chromedriver = 'C:/Users/win10/Desktop/pyt_prac/chromedriver.exe'
    driver = webdriver.Chrome(chromedriver)
    page_num = 1
    #화재정보조사 사이트 접속
    driver.get('http://10.175.105.14/FSSE/lgn/login2.do')
    log_in()
    check_loading()
    pagination()
    choose_condition()
    check_loading()

    #목록 클릭

    # 1페이지
    try:
        page_num = 96
        click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[9]/a/span')
        page_loading()
        #init_page(page_num)
        save_excel('Page',page_num)
    except:
        pass
    try:
        try:
            enter_list_first()
            print("fin first enter")
        except:
            save_excel('Error!!',0)
            time.sleep(100)
            pass
        
        for i in range (1,25):
            try:
                enter_list(i)
            except:
                save_excel('Error!!'+i,0)
                time.sleep(100)
                pass
    except:
        print("Here!2")
        time.sleep(100)
        pass
    try:
        for a in range (10):
            try:
                page_select(7)
                page_num += 1
                save_excel('Page',page_num)
                check_page()

                '''
                click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[9]/a/span')
                page_num = 96
                save_excel('Page',page_num)
                check_page()
                '''
            except:
                save_excel('Error',page_num)
    except:
        save_excel('Error',page_num)
        time.sleep(60)

    print("GGGGGGGGGGGGOOOOOOOOOOOOOOOOODDDDDDDDDDD{}",format(page_num))
