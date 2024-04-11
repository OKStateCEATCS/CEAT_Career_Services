from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from datetime import date
import shutil
import time



#declaring credentials for HireOSUGrads
user_name = "ceatcs@okstate.edu"
pass_word = "hireCEAT2022!"
file_name = "Headshot Check-In (" + date.today().strftime("%m/%d/%Y") + ").xlsx"
id_num="20375533" #input("Please Swipe ID: ")


#This function will allow us to grab the students email from HireOSUGrads
def grab_info():
    driver = webdriver.Chrome()

    #access the url
    driver.get("https://okstate.admin.12twenty.com/students")

    #finding the username input via html and inputting the credentials
    user_input = driver.find_element(By.NAME, 'Username')
    user_input.send_keys(user_name)
    
    #finding the password input via html and inputting the credentials and logging in
    pass_input = driver.find_element(By.NAME, 'Password')
    pass_input.send_keys(pass_word, Keys.ENTER)
    
    #fullscreens the window to find elements
    driver.maximize_window()

    WebDriverWait(driver,20)
    search_bar=driver.find_element(By.XPATH,'//input[contains("Search by Name, Email Address or ID")]')
    search_bar.send_keys(id_num, Keys.ENTER)

    student_file=WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.CLASS_NAME, 'student-list-card-main')))
    student_file.click()

    time.sleep(10)

grab_info()