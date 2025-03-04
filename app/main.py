"""
File:       main.py
Purpose:    This is the main file that begins the execution of the rest of the program
Author:     Joey Borrelli, Software & Training Intern For Catholic Charities of East Tennessee
Datum:      12/17/A.D.2024
"""

import CLI
import data_puller

if __name__ == "__main__":
    print("Welcome to PowerPoint Generator!")
    # start the choose report sequence
    #CLI.choose_report()
    data_puller.pull_data()