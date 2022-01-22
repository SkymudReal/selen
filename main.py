import xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.common.by import By

book = xlwt.Workbook(encoding="utf-8")
rb = xlrd.open_workbook('pochta.xls')# Открываем эксель документ
sheet = rb.sheet_by_index(0)  # Выбираем лист документа
sheet1 = book.add_sheet("Python Sheet 1")

def selen():
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    import time

    #FIRST_NAME = "famousreappointed@rambler.ua"
    #LAST_NAME = "Stn$S38Mli11ZI13"

    browser = webdriver.Chrome(executable_path="C:\chromedriver.exe")#Свое расположение вебдрайвера
    browser.maximize_window()
    browser.get("https://id.rambler.ru/login-20/login?back=http%3A%2F%2Fwww.rambler.ru%2F&rname=main")

    first_name = browser.find_element(By.ID, "login")
    last_name = browser.find_element(By.ID, "password")
    first_name.send_keys(login)
    last_name.send_keys(password)

    select = browser.find_element(By.XPATH, "//*[@id='__next']/div/div/div[2]/div/div/div/div[1]/form/button/span")
    select.click()

    time.sleep(3)

    browser.get("https://mail.rambler.ru/folder/Spam")
    browser.get("https://mail.rambler.ru/folder/Spam/1/?folderName=Spam")

    time.sleep(3)

    select2 = browser.find_element(By.CLASS_NAME, "SpamNotification-button-2a")
    select2.click()

    select3 = browser.find_element(By.XPATH,
                                   "//*[@id='part2']/div/div/div/div/table[3]/tbody/tr[1]/td/table/tbody/tr[4]/td/table/tbody/tr/td/a[1]")
    select3.click()


row_number = sheet.nrows #Количество линий в документе
if row_number > 0:

    for row in range(0,row_number):
        login = []
        password = []
        login.append(str(sheet.row(row)[0]).replace("text:", '').replace("'", ''))
        password.append(str(sheet.row(row)[1]).replace("text:",'').replace("'", ''))
        sheet1.write(row, 0, login)
        sheet1.write(row, 1, password)
        print('\n'.join(login))
        print('\n'.join(password))
        book.save("sdelano.xls")
        selen()
        input('\n'"Нажмите любую клавишу чтобы продолжить... -> "'\n')
else:
    print("Пусто.")