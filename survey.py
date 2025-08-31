import csv
from datetime import datetime
import pickle
import random
import subprocess
import sys
from tkinter import Tk, N, S, E, W, ttk, FALSE, StringVar, Menu


from jinja2 import Environment, FileSystemLoader


from pyisemail import is_email


import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")


day_surgery_address = "D:\\JOHN TILLET\\episode_data\\episodes.csv"

pickle_address = "D:\\JOHN TILLET\\source\\active\\survey\\pickled_date"

recipients = []


def get_pickled_date():
    with open(pickle_address, "rb") as f:
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
        next(reader)
        for episode in reader:
            if len(episode) == 20:
                date = episode[0]
                next_date_dt = datetime.strptime(date, "%d-%m-%Y")
                next_date_dt = next_date_dt.date()
                if next_date_dt > last_date_dt and selector():
                    title = episode[15]
                    first_name = episode[16]
                    last_name = episode[17]
                    email = episode[19]
                    mrn = episode[2]
                    if is_email(email):
                        result = (date, title, first_name, last_name, email, mrn)
                        recipients.append(result)


def make_body(first_name, our_content_id):
    path_to_template = "D:\\JOHN TILLET\\source\\active\\survey"
    loader = FileSystemLoader(path_to_template)
    env = Environment(loader=loader)
    template_name = "body_template.html"
    template = env.get_template(template_name)

    page = template.render(
        first_name=first_name,
        our_content_id=our_content_id
    )

    with open("D:\\JOHN TILLET\\source\\active\\survey\\body.html", "wt") as f:
        f.write(page)

def restart_program():
    subprocess.Popen([sys.executable] + sys.argv)
    sys.exit()

def sender(*args):
    global recipients

    try:
        pat = recipients.pop()
    except IndexError:
        sys.exit(0)
    num_todo = len(recipients)
    num_label["text"] = f"{num_todo} to go."
    if len(recipients) == 0:
        send_but["text"] = "Close"
    else:
        send_but["text"] = "Prepare Next"

    new_last_date_str = pat[0]
    new_last_date_dt = datetime.strptime(new_last_date_str, "%d-%m-%Y")
    new_last_date_dt = new_last_date_dt.date()
    set_new_pickled_date(new_last_date_dt)
    
    # result = (date, title, first_name, last_name, email)
    with open("D:\\JOHN TILLET\\source\\active\\survey\\survey.csv", "at") as f:
        writer = csv.writer(f)
        data = (pat[0], pat[3], pat[5])
        writer.writerow(data)

    logo_path = r"D:\\JOHN TILLET\\source\\active\\survey\\dec_logo.jpg"

    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.Subject = "Patient Survey - Diagnostic Endoscopy Centre"

    email = pat[4]
    mail.To = email
    attachment = mail.Attachments.Add(logo_path)
    CONTENT_ID_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    our_content_id = "my_logo_123"
    attachment.PropertyAccessor.SetProperty(
        CONTENT_ID_PROPERTY, our_content_id)

    first_name = pat[2]
    make_body(first_name, our_content_id)
    html_path = "D:\\JOHN TILLET\\source\\active\\survey\\body.html"
    with open(html_path, 'r', encoding='cp1252') as f:
        html_content = f.read()
    mail.HTMLBody = html_content

    # # Uncomment to actually send the email
    # # mail.Send()

    # Or display it for review before sending
    mail.Display()

def date_picker():
    subprocess.run([sys.executable, "D:\\JOHN TILLET\\source\\active\\survey\\visual_picker.py"])
    # sys.exit(1)
    restart_program()


# start program
last_date_dt = get_pickled_date()
get_patients(last_date_dt)
recipients = list(reversed(recipients))


# set up gui
root = Tk()
root.title("DEC Survey")
root.geometry("350x150+1000+100")
root.option_add("*tearOff", FALSE)

menubar = Menu(root)
root.config(menu=menubar)

menu_startdate = Menu(menubar)
menubar.add_cascade(menu=menu_startdate, label="Set Start Date")
menu_startdate.add_command(label="Set Start Date", command=date_picker)

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
