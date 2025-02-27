import csv
from datetime import datetime
import pickle
import random
import os
from pyisemail import is_email
import sys
import time
from tkinter import Tk, N, S, E, W, ttk, Menu, FALSE, StringVar

import pyautogui as pya
import pyperclip

pya.PAUSE = 0.15
pya.FAILSAFE = True

TITLE_POS = (167, 134)
CLOSE_POS = (780, 95)

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
        with open(day_surgery_address) as f1, open(target_address, "w") as f2:
            reader = csv.reader(f1)
            writer = csv.writer(f2)
            for patient in reader:
                date = patient[0]
                next_date_dt = datetime.strptime(date, "%d-%m-%Y")
                if next_date_dt > last_date_dt and selector():
                    mrn = patient[1]
                    pat_finder(mrn)
                    title, first_name, last_name = front_scrape()
                    email = email_scraper()
                    close_page()
                    if is_email(email):
                        result = (date, title, first_name, last_name, email)
                        print(result)
                        writer.writerow(result)
                        writer.writerow(["", "", "", ""])
            
    except PermissionError:
        pya.alert("Close The EXCEL file!")
        sys.exit()
    return date



def pat_finder(mrn):
    pya.moveTo(100, 300, duration=0.3)
    pya.click()
    pya.hotkey("ctrl", "o")
    pya.hotkey("alt", "b")
    pya.typewrite("f")
    pya.hotkey("alt", "s")
    pya.typewrite(mrn)
    pya.hotkey("alt", "o")

def front_scrape():
    """Scrape name and mrn from blue chip.
    return tiltle, first_name, last_name for meditrust and printname for printed accounts"""

    # TITLE_POS = (200, 134)
    pya.moveTo(TITLE_POS, duration=0.5)
    pya.doubleClick()
    title = pyperclip.copy("na")
    for i in range(4):
        pya.hotkey("ctrl", "c")
        title = pyperclip.paste()
        if title != "na":
            break

    # if title == "na":
    #     pya.alert("Error reading Blue Chip.\nTry again\n?Logged in with AST")


    pya.press("tab")

    first_name = pyperclip.copy("na")
    for i in range(4):
        pya.hotkey("ctrl", "c")
        first_name = pyperclip.paste()
        if first_name != "na":
            break

    if first_name == "na":
        first_name = pya.prompt(
            text="Please enter patient first name", title="First Name", default=""
        )

    pya.press("tab")
    pya.press("tab")
    last_name = pyperclip.copy("na")
    for i in range(4):
        pya.hotkey("ctrl", "c")
        last_name = pyperclip.paste()
        if first_name != "na":
            break
        
    if last_name == "na":
        last_name = pya.prompt(
            text="Please enter patient surname", title="Surame", default=""
        )



    return (title, first_name, last_name)


def close_page():
    """Close patient file with mouse click."""

    pya.moveTo(CLOSE_POS[0], CLOSE_POS[1])
    pya.click()
    # time.sleep(0.25)
    pya.hotkey("alt", "n")
    pya.moveTo(100, 300, duration=0.3)


def email_scraper():
    pya.press("tab", presses=16)
    email = pyperclip.copy("na")
    pya.hotkey("ctrl", "c")
    email = pyperclip.paste()
    email = email.replace(",", "")
    return email

def runner(*args):
    print("Hello from survey!")
    # last_date = "26-02-2025"
    last_date = get_pickled_date()
    last_date_dt = datetime.strptime(last_date, "%d-%m-%Y")
    print(f"Last date:    {last_date_dt}")
    new_last_date = get_patients(last_date_dt)
    set_new_pickled_date(new_last_date)
    close_page()
    os.startfile(target_address)
    time.sleep(3)
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





# if __name__ == "__main__":


