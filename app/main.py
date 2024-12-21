import csv

# Function that selects which parsing algorithm is used depending on the user input
def choose_report():
    valid = False # variable that determines if our input was valid
    while not valid: # will loop again if hit our default case and don't have a good input
        # take in user input
        user_report_choice = input("Which report would you like a powerpoint for?\n1. test\nEnter a choice: ")
        match int(user_report_choice):
            case 1: # test case
                valid = True
                print("TBD")
                with open("reports/test.csv", newline = '') as report:
                    reader = csv.reader(report, delimiter = ' ', quotechar = '|')
                    for row in reader:
                        # figure out how to parse into pptx from here.
                        print(row)
            case 2: # TBD case
                valid = True
                print("TBD")
            case _:
                print("Not a valid choice")
                valid = False

if __name__ == "__main__":
    print("Welcome to PowerPoint Generator")
    choose_report()