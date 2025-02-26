import csv
from datetime import datetime
import pickle
import random
from pyisemail import is_email

import pyautogui as pya
import pyperclip

TITLE_POS = (200, 134)
CLOSE_POS = (774, 96)

day_surgery_address = "D:\\JOHN TILLET\\episode_data\\day_surgery.csv"

email_address = "test@example.com"


def mock_scrape():
    return email_address, "John"


# def get_pickled_date():
#     with open("pickled_date", "rb") as f:
#         return pickle.load(f)


# def set_new_pickled_date(date_str):
#     with open("pickled_date", "ab") as f:
#         pickle.dump(date_str, f)


def selector():
    choice = random.choice([0, 1, 2])
    if choice == 0:
        return True
    return False


def get_patients(last_date_dt):
    with open(day_surgery_address) as f1, open(".\\bccode\\emails.tsv", "w") as f2:
        reader = csv.reader(f1)
        writer = csv.writer(f2, delimiter="\t")
        for patient in reader:
            next_date_dt = datetime.strptime(patient[0], "%d-%m-%Y")
            if next_date_dt > last_date_dt and selector():
                mrn = patient[1]
                pat_finder(mrn)
                title, first_name, last_name = front_scrape()
                email = email_scraper()
                close_page()
                if is_email(email):
                    result = (title, first_name, last_name, email)
                    print(result)
                    writer.writerow(result)
                    writer.writerow(["", "", "", ""])
        # set_new_pickled_date(patient[0])




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
    pya.hotkey("ctrl", "c")
    title = pyperclip.paste()

    # if title == "na":
    #     pya.alert("Error reading Blue Chip.\nTry again\n?Logged in with AST")


    pya.press("tab")

    first_name = pyperclip.copy("na")
    pya.hotkey("ctrl", "c")
    first_name = pyperclip.paste()

    if first_name == "na":
        first_name = pya.prompt(
            text="Please enter patient first name", title="First Name", default=""
        )

    pya.press("tab")
    pya.press("tab")
    last_name = pyperclip.copy("na")
    pya.hotkey("ctrl", "c")
    last_name = pyperclip.paste()
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
    pya.press("tab", presses=14)
    email = pyperclip.copy("na")
    pya.hotkey("ctrl", "c")
    email = pyperclip.paste()
    email = email.replace(",", "")
    return email

# def main():
#     with open("d:\\john tillet\\episode_data\\day_surgery.csv") as file:
#         reader= csv.reader(file)
#         for ep in reader:
#             if ep[0] == date and ep[5] == doctor:
#                 mrn = ep[1]
        
#                 pat_finder(mrn)
#                 pya.moveTo(TITLE_POS, duration=0.5)
#                 front_scrape()
#                 email = email_scraper()
#                 print(email)
#                 close_page()
    


if __name__ == "__main__":
    print("Hello from survey!")
    last_date = "23-02-2025"
    # last_date = get_pickled_date()
    last_date_dt = datetime.strptime(last_date, "%d-%m-%Y")
    print(f"Last date:    {last_date_dt}")
    get_patients(last_date_dt)
    main()
