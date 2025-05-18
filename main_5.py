import argparse
import csv
from datetime import datetime
import pickle
import random
import sys
from tkinter import Tk, N, S, E, W, ttk, FALSE, StringVar

import pymsgbox as pmb
from pyisemail import is_email

import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

parser = argparse.ArgumentParser(description="This is a test script")
parser.add_argument("-t", "--test", action="store_true", help="Run the test")
args = parser.parse_args()
if args.test:
    print("Test mode activated")
    day_surgery_address = "d:\\john tillet\\source\\active\\survey\\test_csv.csv"

else:
    print("No test mode activated")
    day_surgery_address = "D:\\JOHN TILLET\\episode_data\\episodes.csv"


pickle_address = "D:\\JOHN TILLET\\source\\active\\survey\\pickled_date"

recipients = []


def get_pickled_date():
    with open(pickle_address, "rb") as f:
        print(pickle.load(f))
        return pickle.load(f)


def set_new_pickled_date(date_str):
    with open(pickle_address, "wb") as f:
        pickle.dump(date_str, f)


def selector():
    choice = random.choice(
        [
            0,
            1,
            2,
        ]
    )
    if choice == 0:
        return True
    return False


def get_patients(last_date_dt):
    global recipients
    with open(day_surgery_address) as f:
        reader = csv.reader(f)
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
                        recipients.append(result)
    return date


def sender(*args):
    global recipients

    try:
        pat = recipients.pop()
    except IndexError:
        pmb.alert("All done. Remove survey as default signature")
        sys.exit(0)
    num_todo = len(recipients)
    num_label["text"] = f"{num_todo} to go."
    if len(recipients) == 0:
        send_but["text"] = "Close"
    else:
        send_but["text"] = "Prepare Next"

    # pya.click(200, 200)
    # pya.hotkey("ctrl", "n")
    # pya.write(pat[1])
    # pya.press("tab")
    # pya.press("tab")
    # pya.press("tab")
    # pya.write("DEC Survey")
    # pya.press("tab")
    # pya.write(f"Dear {pat[0]},")
    # pya.hotkey("ctrl", "s")
    print(pat)

    first_name = pat[2]
    email = pat[4]
    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.Subject = "DEC Survey"
    mail.Body = f"Dear {first_name},"
    mail.To = email

    # Uncomment to actually send the email
    # mail.Send()

    # Or display it for review before sending
    mail.Display()


last_date = get_pickled_date()
last_date_dt = datetime.strptime(last_date, "%d-%m-%Y")
new_last_date = get_patients(last_date_dt)
recipients = list(reversed(recipients))
set_new_pickled_date(new_last_date)

pmb.alert(f"Greetings. There are {len(recipients)} emails to send since {last_date}")
pmb.alert("Open Outlook and set survey as default signature")


# set up gui
root = Tk()
root.title("DEC Survey")
root.geometry("150x150+800+100")
root.option_add("*tearOff", FALSE)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

num_todo = StringVar()

ttk.Label(mainframe, text="").grid(column=1, row=1, sticky=W)

send_but = ttk.Button(mainframe, text="Start", command=sender)
send_but.grid(column=1, row=2, sticky=E)
# root.bind('<Return>', sender)

num_label = ttk.Label(mainframe, text="")
num_label.grid(column=1, row=6, sticky=W)
num_todo = len(recipients)
num_label["text"] = f"{num_todo} to go"

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)


root.attributes("-topmost", True)
root.mainloop()
