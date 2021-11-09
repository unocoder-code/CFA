import pyautogui
import pyperclip
from selenium import webdriver
import time
import os
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from tkinter import *
import threading

def re_start_writing(종정, adress, name):
    pyautogui.moveTo(941,387)
    pyautogui.click()
    time.sleep(0.1)
    pyautogui.click()
    time.sleep(time1)

    #추가등록 누르기
    pyautogui.moveTo(870,225)
    pyautogui.click()
    time.sleep(1)
    pyautogui.press('esc')

    time.sleep(time2)

    #과제카드명
    pyautogui.moveTo(1357,434)
    pyautogui.click()
    time.sleep(time3)

    pyautogui.moveTo(1208, 371)
    pyautogui.click()
    pyperclip.copy("자체")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')

    time.sleep(0.5)

    pyautogui.moveTo(699,451)
    pyautogui.click()

    time.sleep(0.5)

    pyautogui.moveTo(938,814)
    pyautogui.click()

    time.sleep(1)

    #파일넣기
    pyautogui.moveTo(1345, 503)
    pyautogui.click()

    time.sleep(3)

    pyautogui.moveTo(794,724)
    pyautogui.click()
    pyperclip.copy(adress)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')

    time.sleep(0.5)

    #제목
    pyautogui.moveTo(1350,383)
    pyautogui.click(clicks=3, interval=0.5)
    time.sleep(0.5)
    pyperclip.copy(f'소방시설등 {종정}점검 실시결과 보고서[{name}]')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    #발신기관명
    pyautogui.moveTo(913,733)
    pyautogui.click(clicks= 3, interval=0.5)
    time.sleep(0.5)
    pyperclip.copy(name)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.scroll(-200)
    time.sleep(0.5)

    #열람범위
    pyautogui.moveTo(774,853)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)
    pyautogui.press('down', presses=3, interval=0.5)
    time.sleep(0.5)
    pyautogui.press('enter')

    #담당접수
    pyautogui.moveTo(1303,968)
    pyautogui.click()
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(2)

def log_in():
    WebDriverWait(driver,5).until(EC.presence_of_element_located((By.NAME, 'loginId'))).send_keys('***************')
    driver.find_element_by_name('loginPwd').send_keys('***************')
    driver.execute_script('javascript:sso_login_check()')

def get_site_login():
    global driver
    driver = webdriver.Ie('./IEDriverServer')
    driver.get("http://107.119.0.42:3100/portal/sso/Login/Portal/SSOLogin.jsp")
    driver.implicitly_wait(10)
    log_in()

def sidoportal():
    WebDriverWait(driver,5).until(EC.presence_of_element_located((By.NAME, 'Image_sidoportal'))).click()
    time.sleep(1)
    pyautogui.press('enter')
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

def onnara():
    driver.switch_to.frame('FR_TOP')
    driver.find_element_by_name('CAT_2').click()
    driver.switch_to.default_content()
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

def start_newforming():
    # frame = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, '_MAIN')))
    # driver.switch_to.frame(frame)
    driver.switch_to.frame('_MAIN')
    driver.find_element_by_id('TTL_ENFD5F65507D089BBC32C0BABAD667C0705').click()
    time.sleep(1)

def full_to_name(file_name):
    if file_name.count(".") == 1:
        name = file_name.split(".")[0]
    else:
        for i in range(len(file_name)-1,0,-1):
            if file_name[i] == '.':
                name = file_name[:i]
    return name

def search_file(jakjong):
    folder_address = "C:\\Users\\win10\\Desktop\\"+jakjong
    file_list = os.listdir(folder_address)
    file_list.sort()
    isthising.config(text='~ing')
    for file_fullname in file_list:
        if TF == False:
            break
        file_name = full_to_name(file_fullname) #확장자 없는 이름
        file_rename = folder_address + "\\" + str(f"{jakjong}점검 실시결과 보고서[{file_name}]") + '.pdf'
        os.rename(folder_address + "\\" + file_fullname, file_rename)
        re_start_writing(jakjong, file_rename, file_name)
    isthising.config(text='fin')


# def search_file_jak():
#     search_file('작동기능')
#
# def search_file_jong():
#     search_file('종합정밀')

def thread_jak():
    Thread_jak= threading.Thread(target=search_file, args=['작동기능'])
    Thread_jak.start()

def thread_jong():
    Thread_jong= threading.Thread(target=search_file, args=['종합정밀'])
    Thread_jong.start()

def finish_program():
    global TF
    TF = False
    TrueorFlase.config(text='Stop')

def re_program():
    global TF
    global time1, time2, time3
    TF = True
    TrueorFlase.config(text='Ready')
    time1 = int(time_set1.get())
    time2 = int(time_set2.get())
    time3 = int(time_set3.get())

    print(time1, time2, time3)
def gui_program():
    global TrueorFlase
    global isthising
    global time_set1, time_set2, time_set3

    root = Tk()
    root.title("자체점검 온나라 등록")
    root.geometry("350x400")

    row_num = 0

    site_login = Button(root, width=5, height=1, text="Login", command=get_site_login)
    site_login.grid(row=row_num, column=0, sticky=N + E + W + S, padx=3, pady=3)

    row_num += 1

    작동기능등록 = Button(root, width=16, height=3, text="작동기능 등록", command=thread_jak)
    작동기능등록.grid(row=row_num, column=1, sticky=N + E + W + S, padx=3, pady=3)
    
    종합정밀등록 = Button(root, width=16, height=3, text="종합정밀 등록", command=thread_jong)
    종합정밀등록.grid(row=row_num, column=2, sticky=N + E + W + S, padx=3, pady=3)

    row_num += 1

    isthising = Label(root, text='fin')
    isthising.grid(row=row_num, column=1, columnspan= 2, padx=3,pady=3)

    row_num += 1

    #finish_program

    stop_program = Button(root, width=16, height=3, text="실행 중지", command=finish_program)
    stop_program.grid(row=row_num, column=1, sticky=N + E + W + S, padx=3, pady=3)
    setting = Button(root, width=16, height=3, text="세팅", command=re_program)
    setting.grid(row=row_num, column=2, sticky=N + E + W + S, padx=3, pady=3)

    row_num += 1

    time_cur1 = Label(root, text='time1')
    time_cur1.grid(row=row_num, column=1, padx=3, pady=3)

    time_set1 = Entry(root, width=3)
    time_set1.grid(row=row_num, column=2, padx=3, pady=3)
    time_set1.insert(0, time1)

    row_num += 1

    time_cur2 = Label(root, text='time2')
    time_cur2.grid(row=row_num, column=1, padx=3, pady=3)

    time_set2 = Entry(root, width=3)
    time_set2.grid(row=row_num, column=2, padx=3,pady=3)
    time_set2.insert(0, time2)

    row_num += 1

    time_cur3 = Label(root, text='time3')
    time_cur3.grid(row=row_num, column=1, padx=3, pady=3)

    time_set3 = Entry(root, width=3)
    time_set3.grid(row=row_num, column=2, padx=3, pady=3)
    time_set3.insert(0, time3)

    row_num += 1

    TrueorFlase = Label(root, text='Ready')
    TrueorFlase.grid(row=row_num, column=1, columnspan=2, padx=3, pady=3)

    root.mainloop()


if __name__ == '__main__':
    TF = True
    time1 = 10
    time2 = 12
    time3 = 6
    position = pyautogui.position()
    gui_program()

# # 온나라
# driver.find_element_by_id('G200601121008152682911').click()
# driver.switch_to.window(driver.window_handles[2])
# time.sleep(7)