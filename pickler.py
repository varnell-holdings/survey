import pickle
pickle_address = "D:\\JOHN TILLET\\source\\active\\survey\\pickled_date"


def set_new_pickled_date(last_date):
    with open(pickle_address, "wb") as f:
        pickle.dump(last_date, f)


if __name__ == "__main__":
    last_date = input("Last date 'dd-mm-yyyy':   ")
    set_new_pickled_date(last_date)
        
