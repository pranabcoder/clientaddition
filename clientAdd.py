from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import pandas as pd

executable_path = "./chromedriver"

chrome_options = Options()

chrome_options.add_argument("--incognito")
driver = webdriver.Chrome(executable_path=executable_path, chrome_options=chrome_options)

try:
    read_excel = pd.read_excel('clients.xlsx', sheet_name='Sheet1')
    for i in read_excel.index:
        first_name = str(read_excel['first_name'][i])
        last_name = str(read_excel['last_name'][i])
        email_address = str(read_excel['email_address'][i])
        phone_number = str(read_excel['phone_number'][i])
        person_gender = str(read_excel['gender'][i])
        driver.get('https://portaldev.coachfirst.com/login')
        driver.maximize_window()
        driver.find_element(By.NAME, "email").send_keys("kathleen.a@yopmail.com")
        driver.find_element(By.NAME, "password").send_keys("Test@1234")
        driver.find_element(By.XPATH, "//button[@class='btn  _btn']").click()
        time.sleep(2)
        driver.get("https://portaldev.coachfirst.com/member")
        driver.find_element(By.XPATH, "//*[contains(text(),'Add Client')]").click()
        driver.find_element(By.NAME, "firstName").send_keys(first_name)
        driver.find_element(By.NAME, "lastName").send_keys(last_name)
        driver.find_element(By.NAME, "email").send_keys(email_address)
        driver.find_element(By.NAME, "phone").send_keys(phone_number)
        driver.find_element(By.XPATH, "//div[@class='select__control css-yk16xz-control']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//*[@id='gender']//div[text()='" + person_gender + "']").click()
        driver.find_element(By.XPATH, "//*[@id='addMemberButton']").click()
        time.sleep(2)
        driver.switch_to.alert.accept()
        time.sleep(2)
        driver.find_element(By.XPATH, "//*[@id='navbarDropdownMenuLink']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//a[@class='nav-link px-0']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "(//*[contains(text(),'Logout')])[2]").click()
        print(email_address + " is added successfully")
finally:
    time.sleep(5)
    driver.quit()
