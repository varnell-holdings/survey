import csv
from datetime import datetime
import pickle
import random
from pyisemail import is_email

day_surgery_address = "day_surgery.csv"

email_address = "test@example.com"


def mock_scrape():
    return email_address, "John"


def get_pickled_date():
    with open("pickled_date", "rb") as f:
        return pickle.load(f)


def set_new_pickled_date(date_str):
    with open("pickled_date", "ab") as f:
        pickle.dump(date_str, f)


def selector():
    choice = random.choice([0, 1, 2])
    if choice == 0:
        return True
    return False


def get_patients(last_date_dt):
    with open(day_surgery_address) as f1, open("emails.tsv", "w") as f2:
        reader = csv.reader(f1)
        writer = csv.writer(f2, delimiter="\t")
        for patient in reader:
            next_date_dt = datetime.strptime(patient[0], "%d-%m-%Y")
            if next_date_dt > last_date_dt and selector():
                result = mock_scrape()
                if is_email(result[0]):
                    print(result)
                    writer.writerow(result)
                    writer.writerow(["", ""])
        # set_new_pickled_date(patient[0])


def main():
    print("Hello from survey!")

    last_date = get_pickled_date()
    last_date_dt = datetime.strptime(last_date, "%d-%m-%Y")
    print(f"Last date:    {last_date_dt}")
    get_patients(last_date_dt)


if __name__ == "__main__":
    main()
