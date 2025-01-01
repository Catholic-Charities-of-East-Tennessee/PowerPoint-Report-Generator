"""
File:       GUI.py
Purpose:    This file GUI.py (Graphical User Interface) contains all the methods responsible for I/O (Input Output) and
            the display of content for the user
Author:     Joey Borrelli, Software & Training Intern For Catholic Charities of East Tennessee
Anno:       Anno Domini 2024
"""

import tkinter as tk
import CSV_Interpreter as Interpreter

def start_gui():
    # Create and run window
    root = tk.Tk()
    root.geometry('1300x800')
    root.title('PPTX Generator')
    root.mainloop()
    
def choose_report():
    # take in user input
    report_name = input("\nWhat is the name of the report's csv file located in the reports directory? (ex. \"stats_report\") \nEnter file name: ").strip()
    # if the user typed .csv, this if statement removes that.
    if report_name[-4:] == '.csv':
        report_name = report_name[:-4]
    Interpreter.interpret(report_name)

def get_PowerPoint_Name():
    user_name_choice = input("\nWhat would you like the title slide of the PowerPoint to say?\nEnter title: ")
    return user_name_choice

def get_PowerPoint_SaveName():
    user_name_choice = input("\nWhat would you like to name the pptx file? (ex. \"stats_report\")\nEnter file name: ").strip()
    return user_name_choice

def get_slide_type(title):
    while True:
        user_choice = input("\nWhat type of slide would you like for the data on \"" + str(title) + "\" (1-crosstab, 2-bar graph, 3-pie chart)\nEnter choice: ").strip().lower()

        # if user types in name then change to int
        if user_choice == "crosstab":
            user_choice = 1
        elif user_choice == "bar graph":
            user_choice = 2
        elif user_choice == "pie chart":
            user_choice = 3

        try:
            if int(user_choice) == 1 or user_choice == 2 or user_choice == 3:
                return int(user_choice)
            else:
                print("Invalid choice! Please enter a valid choice...")
        except ValueError:
            print("Invalid choice! Please enter a valid choice...")