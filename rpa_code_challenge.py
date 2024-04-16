from selenium import webdriver
import pandas as pd
import time
import os
from selenium.webdriver.common.by import By


user_name = os.getlogin()

# Read data from the Excel file
file_path = f"C:\\Users\\{user_name}\\PycharmProjects\\RPA Code Challenge\\challenge.xlsx"
challenge_data = pd.read_excel(file_path)

# Open the website
driver = webdriver.Chrome()
driver.get('https://www.rpachallenge.com/')
driver.maximize_window()

start_button = driver.find_element(By.CSS_SELECTOR, 'button.waves-effect.col.s12.m12.l12.btn-large.uiColorButton')
start_button.click()

for _ in range(10):
    for index, row in challenge_data.iterrows():
        # Identify the position of each field
        first_name = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelFirstName"]')
        last_name = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelLastName"]')
        company_name = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelCompanyName"]')
        role_in_company = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelRole"]')
        address = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelAddress"]')
        email = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelEmail"]')
        phone_number = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelPhone"]')

        # Enter data into the correct fields
        first_name.send_keys(row["First Name"])
        last_name.send_keys(row["Last Name "])
        company_name.send_keys(row["Company Name"])
        role_in_company.send_keys(row["Role in Company"])
        address.send_keys(row["Address"])
        email.send_keys(row["Email"])
        phone_number.send_keys(row["Phone Number"])


        submit_button = driver.find_element(By.CSS_SELECTOR, '.btn.uiColorButton')
        submit_button.click()

        # time.sleep(2)
    time.sleep(10)
    driver.quit()
    break
