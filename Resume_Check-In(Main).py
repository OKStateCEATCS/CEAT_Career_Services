from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
import re
import warnings
import os
import time



#declaring credentials for HireOSUGrads
print("Greetings, you will need to input your credentials as well as a link to the event"+"A passcode should be texted to you, please input it into the website.")
user_name = input("Please input your email: ")
pass_word = input("Please input your password: ")
event_link = input("Please input your event link: ")

#this function utilizes selenium to access HireOSUGrads and download the job list
def book_appt():
    driver = webdriver.Chrome()
    
    #searching for the url
    driver.get(event_link)
    
    time.sleep(1)
    
    #finding the username input via html and inputting the credentials
    WebDriverWait(driver,20)
    user_input = driver.find_element(By.NAME, 'username')
    user_input.send_keys(user_name)
    
    #finding the password input via html and inputting the credentials and logging in
    WebDriverWait(driver,20)
    pass_input = driver.find_element(By.NAME, 'password')
    pass_input.send_keys(pass_word, Keys.ENTER)
    
    #fullscrrens the window to find elements
    driver.maximize_window()

    #Finding and waiting for the passcode button to be clickable and then clicking it
    passcd_btn_element=WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Send a passcode")]')))
    passcd_btn_element.click()
    
    #Finding and waiting for the export button to be clickable and then clicking it
    brw_btn_element=WebDriverWait(driver,1000).until(EC.element_to_be_clickable((By.ID, 'dont-trust-browser-button')))
    brw_btn_element.click()
    
    def new_registration():
        rgstr_btn_element=WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"New Registration")]')))
        rgstr_btn_element.click()
        
        email_input = driver.find_element(By.NAME, 'q')
        email_input.send_keys(stdnt_email)
        
        select=Select(driver.find_element(By.NAME, 'q'))
        select.select_by_visible_text(stdnt_cwid)
        
        sub_btn_element=WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"Submit")]')))
        sub_btn_element.click()
    if 
    
    #give time for function to execute
    time.sleep(2)

book_appt()