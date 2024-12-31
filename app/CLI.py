"""
File:       CLI.py
Purpose:    This file CLI.py (Command Line Interface) contains all the methods responsible for I/O (Input Output) to the
            Command line.
Author:     Joey Borrelli, Software & Training Intern For Catholic Charities of East Tennessee
Anno:       Anno Domini 2024
"""

import CSV_Interpreter as Interpreter

def choose_report():
    # take in user input
    report_name = input("\nWhat is the name of the report's csv file located in the reports directory? (ex. \"stats_report\") \nEnter file name: ")
    Interpreter.interpret(report_name)

def get_PowerPoint_Name():
    user_name_choice = input("\nWhat would you like the title slide of the PowerPoint to say?\nEnter name: ")
    return user_name_choice

def get_PowerPoint_SaveName():
    user_name_choice = input("\nWhat would you like to name the pptx file? (ex. \"stats_report\")\nEnter name: ")
    return user_name_choice