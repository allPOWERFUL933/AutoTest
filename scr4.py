import sys
import getopt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import os
import time
from io import BytesIO
import win32clipboard #pywin32
from PIL import Image
import codecs

#python "e:\TEST\scr4.py" -l Gramtest -p 165516sdfs6

def beforetest(argv):
   login = 'Gramtest'
   password = '165516sdfs6'

   try:
    opts, args = getopt.getopt(argv, 'l:p:', ['login=','password=',])
    
   except getopt.GetoptError:
      print('scp.py -l <login> -p <password>')
      sys.exit(2)

   for opt, arg in opts:
      if opt in ('-l', '--login'):
         login = arg
      elif opt in ('-p', '--password'):
         password = arg

   print('логин пароль получены')
   enter(login, password)


def enter(login, password):
   #перейти на яндекс
   driver.get("https://mail.yandex.ru/")
   print('переход на яндекс')
   #войти
   element = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Войти")))
   element.click()
   print('переход к вводу данных аккаунта')
   #логин
   element = wait.until(EC.element_to_be_clickable((By.ID, "passp-field-login")))
   element.send_keys(login)
   print('ввод логина выполнен')
   #далее
   nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
   nextel.click()
   print('переход на ввод пароля')
   #пароль
   element = wait.until(EC.element_to_be_clickable((By.ID, "passp-field-passwd")))
   element.send_keys(password)
   print('ввод пароля выполнен')

   url = driver.current_url
   #далее
   nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
   nextel.click()
   
   wait.until(EC.url_contains("https://mail.yandex.ru/"))

   if driver.current_url != url:
      print('вход выполнен успешно')
   else:
      raise ValueError('ошибка отправки письма')
      pass

      
                                     

def test():
   #написать
   element = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@title='Написать (w, c)']")))
   element.click()
   print('выполнено начало письма')
   #кому                                             
   element = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@class='composeYabbles']")))
   element.send_keys("mr.vaniaschin@yandex.ru")
   print('введен адрес письма')
   #тема
   element = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class='composeTextField ComposeSubject-TextField']")))
   element.send_keys("avtotest")
   print('введена тема письма')
   #письмо
   element = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@role='textbox']")))
   textstr = ("Steven Paul Jobs (/dʒɒbz/; February 24, 1955 – October 5, 2011)"
   "was an American business magnate,  industrial designer, investor,  "
   "and media proprietor. He was the chairman, chief executive officer (CEO), "
   "and co-founder of Apple Inc., the chairman and majority shareholder "
   "of Pixar, a member of The Walt Disney Company's board of directors "
   "following its acquisition of Pixar, and the founder,"
   "chairman, and CEO of NeXT. Jobs is widely recognized as"
   "a pioneer of the personal computer revolution of the 1970s and 1980s,"
   ", along with Apple co-founder Steve Wozniak.")
   element.send_keys(textstr)
   element.send_keys(Keys.RETURN)
   print('введен текст письма')
   #github
   element.send_keys("https://github.com/allPOWERFUL933/AutoTest")
   element.send_keys(Keys.RETURN)
   element.send_keys(Keys.RETURN)
   print('введена ссылка на гитхаб')
   #гиперссылка
   #открыть вставку
   nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='cke_25']")))
   nextel.click()
   #адрес гиперссылки                                          
   nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='href']")))
   nextel.send_keys("https://yandex.ru")
   #текст гиперссылки
   nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='description']")))
   nextel.send_keys("ссылка") 
   #вставить гиперссылку   
   nextel.send_keys(Keys.TAB)
   nextel.send_keys(Keys.RETURN)
   #сбросить выделение ссылки
   element.send_keys(Keys.RIGHT)
   element.send_keys(Keys.RETURN)
   print('выполнена вставка гиперссылки')

   #файл
   nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class='WithUpload-FileInput']")))
   nextel.send_keys(os.path.dirname(sys.argv[0]) + "\\text.PDF")
   print('выполнена загрузка письма')
   
   #картинка
   image = Image.open(os.path.dirname(sys.argv[0]) + "\\pic.jpg")
   output = BytesIO()
   image.convert("RGB").save(output, "BMP")
   data = output.getvalue()[14:]
   output.close()
   win32clipboard.OpenClipboard()
   win32clipboard.EmptyClipboard()
   win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
   win32clipboard.CloseClipboard()
   element.send_keys(Keys.CONTROL,"v")
   print('вставлена картинка')
   #подпись
   element.send_keys("Ваняшин Владислав")
   element.send_keys(Keys.RETURN)
   print('поставлена подпись')

   time.sleep(5)
   #отправить
   nextel = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(concat(' ', @class, ' '), ' ComposeControlPanelButton-Button ')]")))
   nextel.click()
   nextel = None
   nextel = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@class='ComposeDoneScreen-Wrapper']")))

   if nextel != None:
      print('письмо успешно отправлено')
   else:
      raise ValueError('ошибка отправки письма')
      pass



if __name__ == "__main__":
   
   while True:
      try:
         err = False
         pth = os.path.join(os.path.dirname(sys.argv[0]) + "/chromedriver.exe")

         driver = webdriver.Chrome(pth)

         driver.maximize_window()
         wait = WebDriverWait(driver, 30)

         beforetest(sys.argv[1:])
         test()
         break
      except TimeoutException as e:
         print("время ожидания превышено")
         print (e)
         err = True
         break
      except Exception as e:
         print(e)
         err = True
         break
      finally:
         if err == True:
            #сохранить html
            timestr = time.strftime("%Y%m%d-%H%M%S")
            filename = timestr+ "page.html"
            completeName = os.path.join(os.path.dirname(sys.argv[0]) + "\\temp", filename)
            file_object = codecs.open(completeName, "w", "utf-8")
            html = driver.page_source
            file_object.write(html)
            file_object.close()
            #сохранить скрин
            timestr = time.strftime("%Y%m%d-%H%M%S")
            filename = timestr+ "screenshot.png"
            driver.get_screenshot_as_file(os.path.join(os.path.dirname(sys.argv[0]) + "/temp/", filename))
         driver.close()
         driver.quit()
         break