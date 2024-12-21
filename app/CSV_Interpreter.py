import csv

def test(file, cli_instance):
    print("Starting interpreter...")
    try:
        with open("reports/" + file, newline='') as report:
            reader = csv.reader(report, delimiter=',')
            for row in reader:
                # figure out how to parse into pptx from here.
                print(row)
                if row[0] != '':
                    title = row[0]
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