# -*- coding: utf-8 -*-
"""
Created on Mon Apr  7 15:18:32 2025

@author: CEO
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

day_surgery_address = "D:\\JOHN TILLET\\episode_data\\day_surgery.csv"
target_address = "D:\\Jan\\survey\\emails.csv"
pickle_address = "D:\\JOHN TILLET\\source\\active\\survey\\pickled_date"

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
    print("Hello from survey!")
    # last_date = "26-02-2025"
    last_date = get_pickled_date()
    last_date_dt = datetime.strptime(last_date, "%d-%m-%Y")
    print(f"Last date:    {last_date_dt}")
    new_last_date = get_patients(last_date_dt)
    set_new_pickled_date(new_last_date)
    os.startfile(target_address)
    sys.exit(0)


    
# set up gui
root = Tk()
root.title('DEC Survey')
root.geometry('150x150+1300+100')
root.option_add('*tearOff', FALSE)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)


label_date = StringVar()
# date = "26-02-2025"
label_date = get_pickled_date()

ttk.Label(mainframe, text="").grid(column=1, row=1, sticky=W)

but = ttk.Button(mainframe, text='Send!', command=runner)
but.grid(column=1, row=2, sticky=E)
root.bind('<Return>', runner)

ttk.Label(mainframe, text=f"Last Date:  {label_date}").grid(column=1, row=3, sticky=W)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.attributes("-topmost", True)

root.mainloop()