import pickle
pickle_address = "D:\\JOHN TILLET\\source\\active\\survey\\pickled_date"


def get_pickled_date(pickle_address):
    with open(pickle_address, "rb") as f:
        return pickle.load(f)
    
if __name__ == "__main__":
    print(get_pickled_date(pickle_address))