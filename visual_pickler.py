from tkinter import Tk, N, S, E, W, ttk, FALSE
import tkinter as tk
from tkcalendar import Calendar
from datetime import datetime
import pickle


class DatePickerWindow:
    def __init__(self, parent, callback):
        self.parent = parent
        self.callback = callback

        # Create a new top-level window
        self.top = tk.Toplevel(parent)
        self.top.title("Select a Date")

        # Make it modal (will block interaction with the parent window)
        self.top.transient(parent)
        self.top.grab_set()

        # Position it near the parent window
        x = parent.winfo_rootx() + 50
        y = parent.winfo_rooty() + 50
        self.top.geometry(f"+{x}+{y}")

        # Create the calendar widget
        today = datetime.now()
        self.cal = Calendar(
            self.top,
            selectmode="day",
            year=today.year,
            month=today.month,
            day=today.day,
            showweeknumbers=False,
            showothermonthdays=False,
            selectforeground="red",
            foreground="black",
        )
        self.cal.pack(padx=10, pady=10)

        # Add a select button
        select_btn = tk.Button(self.top, text="Select", command=self.select_date)
        select_btn.pack(pady=5)

        # Add a cancel button
        cancel_btn = tk.Button(self.top, text="Cancel", command=self.top.destroy)
        cancel_btn.pack(pady=5)

    def select_date(self):
        # Get the selected date
        selected_date = self.cal.selection_get()

        # Call the callback function with the selected date
        self.callback(selected_date)

        # Close the window
        self.top.destroy()


def on_date_selected(selected_date):
    date_label.config(text=f"Selected date: {selected_date.strftime('%d-%m-%Y')}")
    date_to_pickle = selected_date.strftime("%d-%m-%Y")
    with open("pickled_date", "wb") as f:
        pickle.dump(date_to_pickle, f)
    message_label.grid()
    close_button.grid()


root = Tk()
root.title("Survey Data Picker")
root.geometry("360x280+900+150")
root.option_add("*tearOff", FALSE)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

open_picker_btn = ttk.Button(
    mainframe,
    text="Select Date",
    command=lambda: DatePickerWindow(mainframe, on_date_selected),
)
open_picker_btn.grid(column=1, row=1, sticky=W)

# Label to show the selected date
date_label = ttk.Label(mainframe, text="No date selected")
date_label.grid(column=1, row=2, sticky=W)

message_label = ttk.Label(
    mainframe,
    text="You can now close this program.\nThen close and restart the survey.",
)
message_label.grid(column=1, row=3, sticky=W)

close_button = tk.Button(mainframe, text="Close", command=root.destroy)
close_button.grid(column=1, row=4, sticky=W)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=10)

message_label.grid_remove()
close_button.grid_remove()

root.mainloop()
