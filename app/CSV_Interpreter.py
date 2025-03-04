"""
File:       CSV_Interpreter.py
Purpose:    This file CSV_Interpreter.py is the file with all the methods responsible for taking in the csv data &
            formatting it before sending it off to the PowerPoint Generator. The data for each chart is held between the
            slide_rows & title variables.
Author:     Joey Borrelli, Software & Training Intern For Catholic Charities of East Tennessee
Datum:      12/17/A.D.2024
"""

import csv
import PowerPointGenerator as pptx
import CLI

def interpret(file):
    try:
        # open the report
        with open("reports/" + file + ".csv", newline='') as report:
            # Create instance of the PowerPointGenerator
            pptx_generator = pptx.PowerPointGenerator()

            slide_rows = [] # holds all the rows for a slide
            title = None # holds the title, set to None to identify the first loop

            # set up a csv reader to deal with the delimiter and quote char
            reader = csv.reader(report, delimiter=',', quotechar='"')

            # loop through each row in the file (the row is a list of values)
            for row in reader:

                # if the row begins with text & every value after is empty, then a new chart is being made
                if row[0] != '' and all(element == '' for element in row[1:]):

                    # if this isn't the first iteration, then we need to create a slide with the previous data
                    if title is not None:
                        # create  slide
                        create_slide(slide_rows, title, pptx_generator)
                        # clear the slide_rows array so that we can make a new slide
                        slide_rows = []
                    # set slide/table title value
                    title = row[0]
                # we are in the middle of a chart
                else:
                    slide_rows.append(row)

            # create a slide for the last chart
            create_slide(slide_rows, title, pptx_generator)
            # save presentation
            pptx_generator.save_Presentation()
    except FileNotFoundError:
        print("File not found...")
        CLI.choose_report()
    except PermissionError:
        print("Interpreter failed![PermissionError]...\nPlease contact Joey Borrelli (jborrelli@ccetn.org) with your csv file and terminal output")
    except IndexError:
        print("Interpreter failed![IndexError]...\nPlease contact Joey Borrelli (jborrelli@ccetn.org) with your csv file and terminal output")

def create_slide(chart, title, pptx_generator):
    # Clean up data before sending off to pptx Generator
    # Remove first empty column. if there is an element in the first column, move it to the next column
    chart = [
        [row[0]] + row[2:] if row[0] != '' else row[1:]
        for row in chart
    ]

    # make all rows same length
    # Find chart length
    chart_length = 0
    # loop through each row
    for sRow in chart:
        # loop through each element in a row
        for i in range(len(sRow)):
            if sRow[i] != '' and i > chart_length:
                chart_length = i
    # account for list starting at 0
    chart_length = chart_length + 1
    # set new lines to the proper length
    if len(chart[1]) - chart_length != 0:  # this conditional is because the slide with the most columns will be deleted using this algorithm
        chart = [s_row[:-(len(s_row) - chart_length)] for s_row in chart]

    # remove any instances of 'Count'
    for i in range(len(chart)):  # loop through columns
        for j in range(len(chart[i])):  # loop through rows
            if chart[i][j] == 'Count' or chart[i][j] == 'count':
                chart[i][j] = ''

    # remove any empty rows
    chart = [sublist for sublist in chart if not all(element == '' for element in sublist)]

    # create slide off data
    type_of_slide = CLI.get_slide_type(title)
    match type_of_slide:
        case 1:
            pptx_generator.create_Table_Slide(title, chart, chart_length, len(chart))
        case 2:
            pptx_generator.create_BarChart_slide(title, chart)
        case 3:
            pptx_generator.create_PieChart_Slide(title, chart)