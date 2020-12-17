import sys
import getopt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
from io import BytesIO
import win32clipboard
from PIL import Image
import codecs


# python scr.py -l Gramtest -p 165516sdfs6

def beforetest(argv):
    login = 'Gramtest'
    password = '165516sdfs6'

    try:
        opts, args = getopt.getopt(argv, 'l:p:', ['login=', 'password=', ])

    except getopt.GetoptError:
        print('scp.py -l <login> -p <password>')
        sys.exit(2)

    for opt, arg in opts:
        if opt in ('-l', '--login'):
            login = arg
        elif opt in ('-p', '--password'):
            password = arg

    print('login is ', login)
    print('password is ', password)
    enter(login, password)


def enter(login, password):
    # перейти на яндекс
    driver.get("https://mail.yandex.ru/")
    # войти
    element = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Войти")))
    element.click()
    # логин
    element = wait.until(EC.element_to_be_clickable((By.ID, "passp-field-login")))
    element.send_keys(login)
    # далее
    nextel = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//*[@id='root']/div/div[2]/div[2]/div/div/div[2]/div[3]/div/div/div/div[1]/form/div[3]/button")))
    nextel.click()
    # пароль
    element = wait.until(EC.element_to_be_clickable((By.ID, "passp-field-passwd")))
    element.send_keys(password)
    # далее
    nextel = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//*[@id='root']/div/div[2]/div[2]/div/div/div[2]/div[3]/div/div/div/form/div[3]/button")))
    nextel.click()


def test():
    # написать
    element = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='nb-1']/body/div[2]/div[5]/div/div[3]/div[2]/div[2]/div/div/a")))
    element.click()
    # кому
    element = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                     "//*[@id='nb-1']/body/div[2]/div[9]/div/div/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[1]/div[1]/div/div/div/div/div")))
    element.send_keys("mr.vaniaschin@yandex.ru")
    # тема
    element = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                     "//*[@id='nb-1']/body/div[2]/div[9]/div/div/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[1]/div[3]/div/div/input")))
    element.send_keys("autotest")
    # письмо
    element = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='cke_1_contents']/div")))
    textstr = ("Steven Paul Jobs (/dʒɒbz/; February 24, 1955 – October 5, 2011) was an American business magnate, "
               "industrial designer, investor, and media proprietor. He was the chairman, chief executive officer (CEO), "
               "and co-founder of Apple Inc., the chairman and majority shareholder of Pixar, a member of The Walt Disney "
               "Company's board of directors following its acquisition of Pixar, and the founder, chairman, and CEO of NeXT. "
               "Jobs is widely recognized as a pioneer of the personal computer revolution of the 1970s and 1980s, along with Apple co-founder Steve Wozniak.")
    element.send_keys(textstr)
    element.send_keys(Keys.RETURN)
    # github
    element.send_keys("https://github.com/allPOWERFUL933/AutoTest")
    element.send_keys(Keys.RETURN)
    element.send_keys(Keys.RETURN)
    # гиперссылка
    # открыть вставку
    nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='cke_25']")))
    nextel.click()
    # адрес гиперссылки
    nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='href']")))
    nextel.send_keys("https://yandex.ru")
    # текст гиперссылки
    nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='description']")))
    nextel.send_keys("ссылка")
    # вставить гиперссылку
    nextel.send_keys(Keys.TAB)
    nextel.send_keys(Keys.RETURN)
    # сбросить выделение ссылки
    element.send_keys(Keys.RIGHT)
    element.send_keys(Keys.RETURN)

    # файл
    nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class='WithUpload-FileInput']")))
    nextel.send_keys(os.path.dirname(sys.argv[0]) + "\\text.PDF")

    # картинка
    image = Image.open(os.path.dirname(sys.argv[0]) + "\\pic.jpg")
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()
    element.send_keys(Keys.CONTROL, "v")
    # подпись
    element.send_keys("Ваняшин Владислав")
    element.send_keys(Keys.RETURN)

    # отправить
    nextel = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                    "//*[@id='nb-1']/body/div[2]/div[9]/div/div/div[2]/div/div[2]/div/div[1]/div[1]/div[2]/div/div[1]/div[1]/button")))
    nextel.click()


if __name__ == "__main__":

    while True:
        try:
            pth = os.path.join(os.path.dirname(sys.argv[0]) + "\\chromedriver.exe")
            driver = webdriver.Chrome(pth)

            driver.maximize_window()
            wait = WebDriverWait(driver, 10)

            beforetest(sys.argv[1:])
            test()
            break
        except Exception as e:
            print(e)

            # сохранить html
            timestr = time.strftime("%Y%m%d-%H%M%S")
            filename = timestr + "page.html"
            completeName = os.path.join(os.path.dirname(sys.argv[0]) + "\\temp", filename)
            file_object = codecs.open(completeName, "w", "utf-8")
            html = driver.page_source
            file_object.write(html)
            file_object.close()
            # сохранить скрин
            timestr = time.strftime("%Y%m%d-%H%M%S")
            filename = timestr + "screenshot.png"
            driver.get_screenshot_as_file(os.path.join(os.path.dirname(sys.argv[0]) + "/temp/", filename))

            break
        finally:
            driver.close()
            driver.quit()
            break
