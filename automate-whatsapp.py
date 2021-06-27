import pandas as pd
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import socket


def is_connected():
        try:
                socket.create_connection(('www.google.com', 80))
                return True
        except BaseException:
                print('Try Again')
                is_connected()

def send_whatsapp(driver, name, phone_no, text):
        url = 'https://web.whatsapp.com/send?phone={}&source=&data=#'.format(phone_no)
        driver.get(url)
        sleep(10)

        try:
                input_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
                for ch in text:
                        if ch == "\n":
                            ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
                        else:
                            input_box.send_keys(ch)
                input_box.send_keys(Keys.ENTER)
                sleep(5)
                driver.find_element_by_class_name("web").send_keys(Keys.ENTER)
                print(f"Message sent successfuly\t{name}")
        except Exception:
                print('Invailid phone no ' + str(phone_no))


def main():
    is_connected()
    # Using Chrome web driver and passing path to the chromerdriver.exe file
    driver = webdriver.Chrome(executable_path = r'/mnt/c/Users/USER/chromedriver_win32/chromedriver.exe')
    driver.get('http://web.whatsapp.com')
    sleep(10)

    data = pd.read_excel(r'Participants.xlsx')
    total = len(data.index)

    for i in range(total):
        name = data.at[i, 'Name']
        phoneNo = data.at[i, 'Contact Number']
        if len(str(phoneNo)) == 10:
                phoneNo = '91' + str(phoneNo)
        msg = "Hey " + name + """,
                \n*Add any message* """
        send_whatsapp(driver, name, phoneNo, msg)

if __name__ == '__main__':
    main()
