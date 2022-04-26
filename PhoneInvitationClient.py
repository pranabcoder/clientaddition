from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import time
import pandas as pd

executable_path = "./chromedriver"

chrome_options = Options()

chrome_options.add_argument("--incognito")
driver = webdriver.Chrome(executable_path=executable_path, chrome_options=chrome_options)

try:
    read_excel = pd.read_excel('clients.xlsx', sheet_name='Sheet1')
    for i in read_excel.index:
        email_address = str(read_excel['email_address'][i])
        driver.get('https://portal.coachfirst.com/login')
        driver.maximize_window()
        driver.find_element(By.NAME, "email").send_keys("haleysmith@yopmail.com")
        driver.find_element(By.NAME, "password").send_keys("Haleysmith123")
        driver.find_element(By.XPATH, "//button[@class='btn  _btn']").click()
        time.sleep(2)
        driver.get("https://portal.coachfirst.com/member")
        time.sleep(2)
        driver.find_element(By.XPATH, "//input[@type='search']").send_keys(email_address)
        driver.find_element(By.XPATH, "//*[contains(text(),' Edit')]").click()
        driver.find_element(By.NAME, "isCheckedPhone").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//button[@type='submit' and contains(., 'Update')]").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//*[@id='navbarDropdownMenuLink']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//a[@class='nav-link px-0']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "(//*[contains(text(),'Logout')])[2]").click()
        print(email_address + "Phone invitation send successfully")
finally:
    time.sleep(5)
    driver.quit()
