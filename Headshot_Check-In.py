from openpyxl import load_workbook
from datetime import date
from tkinter import *
import tkinter as tk
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
    def close_iwin(e):
        global iso
        iso=iso_var.get()
        iwin.destroy()
    
    iwin=tk.Tk()
    iwin.rowconfigure(0, weight=2)
    iwin.rowconfigure(3, weight=2)
    iwin.columnconfigure(0, weight=2)
    iwin.columnconfigure(3, weight=2)
    iwin.title("Info Search")
    l = Label(iwin, text="Please Swipe Your ID.")
    l.grid(row=1, column=0, columnspan=4)
    global iso_var
    iso_var=tk.StringVar()
    i=Entry(iwin,textvariable=iso_var)
    i.focus_set()
    i.grid(row=2, column=2)
    iwin.bind('<Return>', close_iwin)
    
    iwin.state("zoomed")
    iwin.mainloop()
    
    #implementing a text based kill switch to end the program
    if(iso=="exit"):
        print("Exiting Program")
        time.sleep(0.5)
        exit()
    #Declaring book and ignoring a vanity related warning to clean up the output
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
        book = load_workbook(r''+year_dir+file_name) 
    ws=book['Export']
    
    current_row=0

    #If we find the iso, we confirm their info, otherwise we search using email
    trigger=0

    for column in ws.iter_cols():
        column_name=column[0].value
        if column_name == "Student Email":
            global email_column
            email_column=column
        elif column_name == "Student Swipe Number":
            global iso_column
            iso_column=column

    for cell in iso_column:
        current_row+=1
        temp_val = cell.value
        if (temp_val == iso):
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
            email_cell=ws.cell(row=current_row, column=iso_column)
            email=email_cell.value
            info_confirmation()               
    if (trigger==0):
        print('Unable to Locate Information.')
        search_email()
      
      
      
def search_email():
    def submit():
        #assigning the global email param to the inputted email string
        global email
        email=email_var.get()
        ewin.destroy()
        #Declaring book and ignoring a vanity related warning to clean up the output
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
            book = load_workbook(r''+year_dir+file_name) 
        ws=book['Export']
    
        current_row=0
    
        #If we find it, we confirm their info, otherwise we get email
        for cell in ws['F']:
            current_row+=1
            temp_val = cell.value
            if temp_val == email:
            #Retrieving Student's First Name
                global f_name
                fn_cell=ws.cell(row=current_row, column=2)
                f_name=fn_cell.value
            #Retrieving Student's Last Name       
                global l_name
                ln_cell=ws.cell(row=current_row, column=3)
                l_name=ln_cell.value   
        info_confirmation()
    
    ewin=tk.Tk()
    ewin.rowconfigure(0, weight=2)
    ewin.rowconfigure(4, weight=2)
    ewin.columnconfigure(0, weight=2)
    ewin.columnconfigure(3, weight=2)
    email_var=tk.StringVar()
    ewin.title("Email Confirmation")
    
    #creating a label for email prompt
    eil=Label(ewin,text="Please input your email below:")
    eil.grid(row=1, column=0, columnspan=4)
    
    #creating label for email input
    el=Label(ewin,text="Email:")
    el.grid(row=2, column =1)
    
    #creating entry box for email
    e=Entry(ewin,textvariable=email_var)
    e.grid(row=2, column=2)
    
    #creating submit button
    sub_btn=tk.Button(ewin, text="Submit", command=submit)
    sub_btn.grid(row=3, column=0, columnspan=4)
    
    ewin.state("zoomed")
    ewin.mainloop()
    

        
  
def info_confirmation():
    global l_name
    global f_name
    global email
    
    def confirmed():
        win.destroy()
        send_email()
      
    win=tk.Tk()
    win.rowconfigure(0, weight=2)
    win.rowconfigure(3, weight=2)
    win.columnconfigure(0, weight=2)
    win.columnconfigure(3, weight=2)
    win.title("Info Conformation")
    l = Label(win, text="Is this your information?\n"+f_name+" "+l_name+"\n"+email)
    l.grid(row=1, column=0, columnspan=4)
    yes_btn=tk.Button(win, text="Yes", command=confirmed)
    yes_btn.grid(row=2, column=1)
    no_btn=tk.Button(win, text="No", command=search_email)
    no_btn.grid(row=2, column=2)
    
    win.state("zoomed")
    win.mainloop()
    
    
    
    
    




          
def send_email():
    #for Ethan
    
    #Put vv this vv at the end of your code
    restart()




   
def restart():
    os.system('cls')
    search_ISO()





format_folder()