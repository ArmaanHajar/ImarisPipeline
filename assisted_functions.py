import tkinter as tk
from tkinter import filedialog

def select_folder(message: str):
    root = tk.Tk()
    root.withdraw()

    folder_path = filedialog.askdirectory(title=message)

    return folder_path

def float_input(message: str):
    while True:
        x = input(message)
        try:
            return float(x)
        except ValueError:
            print("Please enter a valid number! (ex. 1 or 1.1)")

def y_or_n(char: str):
    while char.lower() not in ['y', 'n']:
        char = input("Invalid input. Please enter 'y' or 'n': ")
    return char.lower() == 'y'