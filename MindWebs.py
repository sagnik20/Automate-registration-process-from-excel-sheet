import xlrd
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time


sagnk = xlrd.open_workbook('''Link to the Excel file goes here. only local address''')

sheet = sagnk.sheet_by_index(0)
sheet.cell_value(0, 0)

for i in range(sheet.nrows):
    print(sheet.cell_value(i, 0))

for i in range(sheet.nrows):
    try:
        driver = webdriver.Firefox()
        driver.get("http://mindwebs.org/events/techtronista/eventRegister.php")
        time.sleep(5)

        user_name_elem = driver.find_element_by_xpath("//input[@name='name']")
        user_name_elem.clear()
        user_name_elem.send_keys(sheet.cell_value(i,0))

        user_email = driver.find_element_by_xpath("//input[@name='email']")
        user_email.clear()
        user_email.send_keys(sheet.cell_value(i,1))

        user_phone = driver.find_element_by_xpath("//input[@name='phone']")
        user_phone.clear()
        user_phone.send_keys(sheet.cell_value(i,2))

        user_add = driver.find_element_by_xpath("//input[@name='address']")
        user_add.clear()
        user_add.send_keys(sheet.cell_value(i,3))

        user_prof = driver.find_element_by_xpath("//span[@title='Select Your Profession']")
        user_prof.click()
        abc = driver.find_element_by_xpath("//input[@class='select2-search__field']")
        abc.send_keys("Student")
        abc.send_keys(Keys.RETURN)

        user_wadd = driver.find_element_by_xpath("//input[@name='work_at']")
        user_wadd.clear()
        user_wadd.send_keys(sheet.cell_value(i,4))

        user_day = driver.find_element_by_xpath("//span[@title='Select Your Event Days']")
        user_day.click()
        defi = driver.find_element_by_xpath("//input[@class='select2-search__field']")
        defi.send_keys("Both Day 1 and 2")
        defi.send_keys(Keys.RETURN)

        user_reg = driver.find_element_by_xpath("//input[@name='register']")
        user_reg.click()

        time.sleep(5)
        driver.close()
    except Exception:
        print("There is some problem in running the code. Please contact SAGNIK CHATTERJEE for help.")