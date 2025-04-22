# -*- coding: utf-8 -*-
"""
Created on Wed Apr  9 07:29:40 2025

@author: John2
"""

import csv
from datetime import datetime
import os
import pickle
import random
import sys
from tkinter import Tk, N, S, E, W, ttk, Menu, FALSE, StringVar

import pyautogui as pya
from pyisemail import is_email
import pyperclip

pya.PAUSE = 0.3♥
pya.FAILSAFE = True

day_surgery_address = "D:\\JOHN TILLET\\episode_data\\day_surgery.csv"
target_address = "D:\\Jan\\survey\\emails.csv"
pickle_address = "D:\\JOHN TILLET\\source\\active\\survey\\pickled_date"

test_address = "d:\\john tillet\\source\\active\\survey\\test_csv.csv"


global recipients

def get_pickled_date():
    with open(pickle_address, "rb") as f:
        return pickle.load(f)

def set_new_pickled_date(date_str):
    with open(pickle_address, "wb") as f:
        pickle.dump(date_str, f)


def selector():
    choice = random.choice([0, 1, ])
    if choice == 0:
        return True
    return False

def get_patients(last_date_dt):
    try:
        with open(day_surgery_address) as f, open(target_address, "w") as f2:
            reader = csv.reader(f)
            writer = csv.writer(f2)
            for episode in reader:
                if len(episode) == 20:
                    date = episode[0]
                    next_date_dt = datetime.strptime(date, "%d-%m-%Y")
                    if next_date_dt > last_date_dt and selector():
                        title = episode[15]
                        first_name = episode[16]
                        last_name = episode[17]
                        email = episode[19]
                        if is_email(email):
                            result = (date, title, first_name, last_name, email)
                            print(result)
                            writer.writerow(result)
                            writer.writerow(["", "", "", ""])
    except PermissionError:
        pya.alert("Close The EXCEL file!")
        sys.exit()
        return date


def runner(*args):
    global recipients
    print("Hello from survey!")
    recipients = [["John", "tillett1957@gmail.com"], ["Fred", "john.lamia@gmail.com"]]
    # last_date = "26-02-2025"
    # last_date = get_pickled_date()
    # last_date_dt = datetime.strptime(last_date, "%d-%m-%Y")
    # print(f"Last date:    {last_date_dt}")
    # new_last_date = get_patients(last_date_dt)
    # set_new_pickled_date(new_last_date)
    # os.startfile(target_address)
    # sys.exit(0)
    col_but.grid_remove()
    send_but.grid()
    num_label.grid()
    num_todo = len(recipients)
    num_label["text"] = f"{num_todo} to go."
    pya.alert("Open Outlook and set survey as default signature")

def sender(*args):
    global recipients
    
    try:
        pat = recipients.pop()
    except IndexError:
        print("all done")
        pya.alert("Remove survey as default signature")
        sys.exit(0)
    num_todo = len(recipients)
    num_label["text"] = f"{num_todo} to go."
    if len(recipients) == 0:
        send_but["text"] = "Close"
    else:
        send_but["text"] = "Next"
        
    pya.click(200, 200)
    pya.hotkey("ctrl", "n")
    pya.write(pat[1])
    pya.press("tab")
    pya.press("tab")
    pya.press("tab")
    pya.write("DEC Survey")
    pya.press("tab")
    pya.write(f"Dear {pat[0]},")
    pya.hotkey("ctrl", "s")
    print(pat)


    
# set up gui
root = Tk()
root.title('DEC Survey')
root.geometry('250x250+800+100')
root.option_add('*tearOff', FALSE)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)


label_date = StringVar()
num_todo = StringVar()
# date = "26-02-2025"
label_date = get_pickled_date()

ttk.Label(mainframe, text="").grid(column=1, row=1, sticky=W)

col_but = ttk.Button(mainframe, text='Collect emails', command=runner)
col_but.grid(column=1, row=2, sticky=E)
# root.bind('<Return>', runner)

ttk.Label(mainframe, text=f"Last Date:  {label_date}").grid(column=1, row=3, sticky=W)

ttk.Label(mainframe, text="").grid(column=1, row=4, sticky=W)

send_but = ttk.Button(mainframe, text='Send Mail', command=sender)
send_but.grid(column=1, row=5, sticky=E)
# root.bind('<Return>', sender)
num_label = ttk.Label(mainframe, text="")
num_label.grid(column=1, row=6, sticky=W)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

send_but.grid_remove()
num_label.grid_remove()

root.attributes("-topmost", True)
root.mainloop()