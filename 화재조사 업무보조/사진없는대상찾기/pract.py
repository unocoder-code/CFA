from json import load, loads
from re import search
from selenium import webdriver
from bs4 import BeautifulSoup
import openpyxl
import time


def check_file():
    html_raw = driver.page_source
    html = BeautifulSoup(html_raw, 'html.parser')
    for i in range(1,5):
    #for i in range(4,0,-1):
        file_code = html.find(id="tr_file-{}".format(i))
        file_attrs = file_code.attrs
        print(file_attrs)
        if 'class' in file_attrs:
            print('inserted')
        else:
            print('Not image')
            Append_Not_image_list(html)
            break

#사진 없을 경우 엑셀에 목록 추가
def Append_Not_image_list(sorce):
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
    send_key_id('_easyui_textbox_input1','0105108374311')
    send_key_id('_easyui_textbox_input2','0000')
    time.sleep(0.5)
    click_mouse_id("btnLogin")
    time.sleep(1)

def check_loading():
    html_raw = driver.page_source
    html = BeautifulSoup(html_raw, 'html.parser')
    loadingcode = html.find(id="fsseProgress")
    load_attrs = loadingcode.attrs
    if 'in' in load_attrs['class']:
        print('~~~loading~~~')
        time.sleep(0.5)
        check_loading()
    else:
        print('End loading')

def page_loading():
    html_raw = driver.page_source
    html = BeautifulSoup(html_raw, 'html.parser')
    page_code = html.find(class_="datagrid-mask")
    if page_code == None:
        print('ready!')
    else:
        page_loading()

def next_page():
    click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[8]/a/span')


def pagination():
    click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[1]/select/option[4]')

chromedriver = 'C:/Users/win10/Desktop/pyt_prac/chromedriver.exe'
driver = webdriver.Chrome(chromedriver)

#화재정보조사 사이트 접속
driver.get('http://10.175.105.14/FSSE/lgn/login2.do')
log_in()
check_loading()
try:
    print("Hi hello")
    click_mouse_xpath('//*[@id="contentList"]dffaefdiv[2]/table/tbody/tr/td[6]/a[2]/span/span')
    print("error!")
except:
    print("error123!")
    click_mouse_xpath('//*[@id="contentList"]dffaefdiv[2]/table/tbody/tr/td[6]/a[2]/span/span')
    print("error!")
    next_page()
time.sleep(10000)
'''
click_mouse_xpath('//*[@id="contentList"]/div[2]/div/div[2]/table/tbody/tr/td[6]/a[2]/span/span')
page_loading()
'''