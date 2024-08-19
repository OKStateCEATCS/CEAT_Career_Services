from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from datetime import date
import os
import re
import warnings
import time







#declaring query file name
file_name = "CEAT Student Data ({}).xlsx"
if 1<=int(date.today().strftime("%m"))<=12:
    file_name = file_name.format("Spring, "+date.today().strftime("%Y"))
#elif 6<=int(date.today().strftime("%m"))<=7:
#    file_name = file_name.format("Summer, "+date.today().strftime("%Y"))
#else:
#    file_name = file_name.format("Fall, "+date.today().strftime("%Y"))







#Finding and Formatting the Query Folder
def format_folder():
    outer_dir="K:\\studentservices\\crsvc_sh\\Headshot Check In\\"
    inner_dir="{} Queries\\"
    inner_dir=inner_dir.format(date.today().strftime("%Y"))
    #this is combining our two path parts into one and setting that as our path
    path=os.path.join(outer_dir,inner_dir)
    global year_dir
    year_dir=path
    #we are either making our folder or it already exists and we go to grab_excel
    try:
        os.mkdir(path)
        search_ISO()
    except FileExistsError:
        search_ISO()
    except FileNotFoundError:
        print("The generated path is invalid; check the directories for changes.")
   
   
   
   
   
        
#This function will allow us to grab the students info from the excel sheet query using their ISO Number
def search_ISO():
    #Getting ISO number from card
    iso=input("Please Swipe ID: ")
    if(iso=="exit"):
        print("Terminal Killed")
        time.sleep(1)
        exit()
    #Declaring book and ignoring a vanity related warning to clean up output
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
        book = load_workbook(r''+year_dir+file_name) 
    ws=book['Export']
    
    current_row=0

    #If we find it, we confirm their info, otherwise we search using ISO
    trigger=0
    for cell in ws['E']:
        current_row+=1
        temp_val = cell.value
        if temp_val == iso:
            trigger=1
        #Retrieving Student's First Name
            global f_name
            fn_cell=ws.cell(row=current_row, column=2)
            f_name=fn_cell.value
        #Retrieving Student's Last Name       
            global l_name
            ln_cell=ws.cell(row=current_row, column=3)
            l_name=ln_cell.value
        #Retrieving Student's CWID
            global cwid
            cwid_cell=ws.cell(row=current_row, column=4)
            cwid=cwid_cell.value    
        #Retrieving Student's Email
            global email
            email_cell=ws.cell(row=current_row, column=6)
            email=email_cell.value
            info_confirmation()
                   
    if (trigger==0):
        print('Unable to Locate Information.')
        search_CWID()
      
      
      
      
      
          
def search_CWID():
    global cwid
    cwid=input("Please Input Your CWID (Include the A): ")
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
        book = load_workbook(r''+year_dir+file_name) 
    ws=book['Export']
    
    current_row=0
    
    #If we find it, we confirm their info, otherwise we get email
    trigger=0
    for cell in ws['D']:
        current_row+=1
        temp_val = cell.value
        if temp_val == cwid:
            trigger=1
        #Retrieving Student's First Name
            global f_name
            fn_cell=ws.cell(row=current_row, column=2)
            f_name=fn_cell.value
        #Retrieving Student's Last Name       
            global l_name
            ln_cell=ws.cell(row=current_row, column=3)
            l_name=ln_cell.value   
        #Retrieving Student's Email    
            global email
            email_cell=ws.cell(row=current_row, column=6)
            email=email_cell.value
            info_confirmation()
    if trigger==0:
        print('Unable to Locate Information.')
        search_email()
        
        
        
        
        
def search_email():
    global email
    email=input("Please Input Your email: ")
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
        book = load_workbook(r''+year_dir+file_name) 
    ws=book['Export']
    
    current_row=0
    
    #If we find it, we confirm their info, otherwise we get email
    trigger=0
    for cell in ws['F']:
        current_row+=1
        temp_val = cell.value
        if temp_val == email:
            trigger=1
        #Retrieving Student's First Name
            global f_name
            fn_cell=ws.cell(row=current_row, column=2)
            f_name=fn_cell.value
        #Retrieving Student's Last Name       
            global l_name
            ln_cell=ws.cell(row=current_row, column=3)
            l_name=ln_cell.value   
        #Retrieving Student's Email    
            global cwid
            cwid_cell=ws.cell(row=current_row, column=4)
            cwid=cwid_cell.value 
            info_confirmation()
    if trigger==0:
        print('Unable to Locate Information.')
        email_confirmation()




def email_confirmation():
    global email
    confirmation=0
    print("Is this your email?\n"+email)
    confirmation=input("Input 0 For No, or 1 For Yes: ")
    if int(confirmation)==1:
        send_email()
    else:
        email=input("Please input your email: ")
        email_confirmation()
        
        
        
        
def info_confirmation():
    global email
    confirmation=0
    print("Is this your information?\n"+l_name+", "+f_name+"\n"+email+"\n"+cwid)
    confirmation=input("Input 0 For No, or 1 For Yes: ")
    if int(confirmation)==1:
        send_email()
    else:
        email=input("Please input your email: ")
        send_email()




          
def send_email():
    print("For Ethan")
    restart()




   
def restart():
    os.system('cls')
    search_ISO()





format_folder()