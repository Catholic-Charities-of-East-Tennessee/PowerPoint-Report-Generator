import csv
import PowerPointGenerator as pptx

def test(file, cli_instance):
    print("Starting interpreter...")
    try:
        # open the report
        with open("reports/" + file + ".csv", newline='') as report:

            # Create instance of the PowerPointGenerator
            pptx.PowerPointGenerator(cli_instance)

            reader = csv.reader(report, delimiter=',', quotechar='"')
            # loop through each row in the file (the row is a list of values)
            for row in reader:
                print(row)

                # if the row begins with text then a new chart is being made
                if row[0] != '':
                    # the first
                    title = row[0]

                # we are in the middle of a chart
                #else:
                    #print("TBD")
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