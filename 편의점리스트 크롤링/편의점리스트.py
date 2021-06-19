from json import load, loads
from re import search
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import *
import time

def save_excel(data):
    wb = openpyxl.load_workbook('편의점목록.xlsx')
    sheet1 = wb.active
    sheet1.append([data[5],data[6],data[7],data[0],data[2],data[4],data[3],data[1]])
    wb.save('편의점목록.xlsx')

def get_data():
    data = []
    driver.switch_to.frame(driver.find_element_by_id("tgdc"))
    driver.switch_to.frame(driver.find_element_by_id("map"))
    for i in range(1,6):
        data.append(driver.find_element_by_xpath('/html/body/table/tbody/tr[1]/td/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[{}]/font'.format(i)).text)
    driver.switch_to.default_content()
    return data

def switch_frame(frameid):
    driver.switch_to.frame(driver.find_element_by_id("tgdc"))
    driver.switch_to.frame(driver.find_element_by_id("IFMfamily"))
    if frameid != 0:
        driver.switch_to.frame(driver.find_element_by_id(frameid))

def list_len(frame,xpath):
    switch_frame(frame)
    num = len(driver.find_element_by_xpath(xpath).text.split())
    driver.switch_to.default_content()
    return num

if __name__ == '__main__':
    chromedriver = './chromedriver.exe'
    driver = webdriver.Chrome(chromedriver)
    driver.get('https://www.famiport.com.tw/Web_Famiport/page/ShopQuery.aspx#')
    driver.implicitly_wait(10)

    #시 선택
    for si in range (2,list_len(0,'/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/select[1]')+1):
        switch_frame(0)
        slect_si = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/select[1]/option[{}]'.format(si))
        si_name = slect_si.text
        slect_si.click()
        driver.switch_to.default_content()
        
        #군 선택
        for gun in range (2,list_len(0,'/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/select[2]')+1):
            switch_frame(0)
            slect_gun = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/select[2]/option[{}]'.format(gun))
            gun_name = slect_gun.text
            slect_gun.click()
            driver.switch_to.default_content()
            
            #구 선택
            for gu in range (2,list_len('street','//*[@id="form1"]/select')-10):
                switch_frame('street')
                slect_gu = driver.find_element_by_xpath('//*[@id="form1"]/select/option[{}]'.format(gu))
                gu_name = slect_gu.text
                slect_gu.click()
                driver.switch_to.default_content()
                time.sleep(0.1)

                #편의점 선택
                print(list_len('store','/html/body/form/select'))
                connum = int((list_len('store','/html/body/form/select')-1)/2+2)
                print(connum)
                for con in range (2,connum):
                    switch_frame('store')
                    driver.find_element_by_xpath('/html/body/form/select/option[{}]'.format(con)).click()
                    driver.switch_to.default_content()

                    #검색 버튼 클릭
                    switch_frame(0)
                    driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div/a/img').click()
                    driver.switch_to.default_content()
                    result = get_data()
                    result.extend([si_name,gun_name,gu_name])
                    print(result)
                    save_excel(result)
