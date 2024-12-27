import csv
import PowerPointGenerator as pptx

def test(file, cli_instance):
    print("Starting interpreter...")
    try:
        # open the report
        with open("reports/" + file + ".csv", newline='') as report:

            # Create instance of the PowerPointGenerator
            pptx_generator = pptx.PowerPointGenerator(cli_instance)
            # holds all the rows for a slide
            slide_rows = []
            title = None

            reader = csv.reader(report, delimiter=',', quotechar='"')
            # loop through each row in the file (the row is a list of values)
            for row in reader:
                # if the row begins with text then a new chart is being made
                if row[0] != '':
                    # if this isn't the first time, then we need to create a slide with the date
                    if title is not None:
                        # Clean up data before sending off to pptx Generator
                        # remove first empty value
                        slide_rows = [row[1:] for row in slide_rows]

                        # make all rows same length


                        # create slide off data
                        pptx_generator.create_Table_Slide(title, slide_rows)
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