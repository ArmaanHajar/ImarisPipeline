# This file contains functions that are used to assist the user in inputting data.

def float_input(message: str):
    while True:
        x = input(message)
        try:
            return float(x)
        except ValueError:
            print("Please enter a valid number! (ex. 1 or 1.1)")

def y_or_n(message: str):
    while True:
        char = input(message + " (y/n): ").strip().lower()
        if char in ['y', 'n']:
            return char == 'y'
        else:
            print("Invalid input. Please enter 'y' or 'n'.")