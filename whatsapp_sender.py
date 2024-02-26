from openpyxl import Workbook,load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

def run_webdriver(excel_file_path, massage):
    wb = load_workbook(excel_file_path)
    ws = wb.active

    text = massage
    web = webdriver.Chrome()
    flag=1
    for x in range(1, ws.max_row + 1):
        phone = ws.cell(row=x, column=1).value
        phone = str(phone)
        phone = phone.replace("-", "")
        if len(phone) >= 10:
            # phone=phone[:10]
            print(x, phone, len(phone))
            link = ('https://wa.me/' + '972' + phone + '?text=' + text)
            web.get(link)
            web.implicitly_wait(3)
            try:
                web.find_element(By.XPATH,
                                 '/html/body/div[1]/div[1]/div[2]/div/section/div/div/div/div[2]/div[1]/a/span').click()
            except:
                web.implicitly_wait(1)
                continue
            web.implicitly_wait(3)
            web.find_element(By.XPATH,
                             '/html/body/div[1]/div[1]/div[2]/div/section/div/div/div/div[3]/div/div/h4[2]/a/span').click()
            if flag == 1:
                flag=0
                web.implicitly_wait(40)
            else:
                web.implicitly_wait(3)

            try:
                sleep(4)
                web.find_element(By.XPATH,
                                 '/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(massage)
            except:
                web.implicitly_wait(1)

            try:
                sleep(4)
                web.find_element(By.XPATH,
                                 '/html/body/div[1]/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]').click()
            except:
                web.implicitly_wait(1)
            sleep(3)
    sleep(10)
    web.close()

#run_webdriver("Book1.xlsx","HI")