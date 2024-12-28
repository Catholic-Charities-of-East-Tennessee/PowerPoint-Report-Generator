import csv
import PowerPointGenerator as pptx

def test(file, cli_instance):
    print("Starting interpreter...")
    try:
        # open the report
        with open("reports/" + file + ".csv", newline='') as report:
            # Create instance of the PowerPointGenerator
            pptx_generator = pptx.PowerPointGenerator(cli_instance)

            slide_rows = [] # holds all the rows for a slide
            title = None # holds the title, set to None to identify the first loop

            # set up a csv reader to deal with the delimiter and quote char
            reader = csv.reader(report, delimiter=',', quotechar='"')

            # loop through each row in the file (the row is a list of values)
            for row in reader:
                # if the row begins with text & every value after is empty, then a new chart is being made
                if row[0] != '' and all(element == '' for element in row[1:]):
                    # if this isn't the first time, then we need to create a slide with the date
                    if title is not None:
                        # Clean up data before sending off to pptx Generator
                        print("\nDefault slide_rows")
                        for r in slide_rows:
                            print(r)

                        # Remove first empty column. if there is an element in the first column, move it to the next column
                        slide_rows = [
                            [row[0]] + row[2:] if row[0] != '' else row[1:]
                            for row in slide_rows
                        ]
                        print("\nremove first empty column")
                        for r in slide_rows:
                            print(r)

                        # make all rows same length
                        # Find chart length
                        chart_length = 0
                        # loop through each row
                        for sRow in slide_rows:
                            # loop through each element in a row
                            for i in range(len(sRow)):
                                if sRow[i] != '' and i > chart_length:
                                    chart_length = i
                        # account for list starting at 0
                        chart_length = chart_length + 1
                        # set new lines to the proper length
                        slide_rows = [s_row[:-(len(sRow) - chart_length)] for s_row in slide_rows]
                        print("\nmake all rows same length")
                        for r in slide_rows:
                            print(r)

                        # remove any instances of 'Count'
                        for i in range(len(slide_rows)):  # loop through columns
                            for j in range(len(slide_rows[i])):  # loop through rows
                                if slide_rows[i][j] == 'Count':
                                    slide_rows[i][j] = ''
                        print("\nremove instances of count")
                        for r in slide_rows:
                            print(r)

                        # remove any empty rows
                        slide_rows = [sublist for sublist in slide_rows if not all(element == '' for element in sublist)]
                        print("\nremove emtpy rows")
                        for r in slide_rows:
                            print(r)

                        # create slide off data
                        pptx_generator.create_Table_Slide(title, slide_rows, chart_length, len(slide_rows))

                        # clear rows
                        slide_rows = []
                    # set slide/table title value
                    title = row[0]
                # we are in the middle of a chart
                else:
                    slide_rows.append(row)
            # save presentation
            pptx_generator.save_Presentation()
    except FileNotFoundError:
        print("File not found...\n")
        cli_instance.valid = False
    except PermissionError:
        print("Interpreter failed![PermissionError]...\nPlease contact Joey Borrelli (jborrelli@ccetn.org) with your csv file and terminal output")
    except IndexError:
        print("Interpreter failed![IndexError]...\nPlease contact Joey Borrelli (jborrelli@ccetn.org) with your csv file and terminal output")


def stats_report(file, cli_instance):
    print("Starting interpreter...")

def stats_report_with_hope_kitchen(file, cli_instance):
    print("Starting interpreter...")