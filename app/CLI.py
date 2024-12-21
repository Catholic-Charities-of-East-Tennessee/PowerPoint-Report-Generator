import CSV_Interpreter as interpreter

class CLI:
    def __init__(self):
        # variable that determines if our input was valid
        self._valid = False

    @property
    def valid(self):
        return self._valid

    @valid.setter
    def valid(self, state):
        self._valid = state

    def choose_report(self, interface_instance):
        # will loop again if hit our default case and don't have a good input
        while not self._valid:
            # take in user input
            user_report_choice = input(
                "Which report would you like a powerpoint for?\n1. test\n2. TBD\nEnter a choice: ")
            report_name = input("What is the name of the csv file? (ex. \"stats_report.csv\") \nEnter name: ")

            # match case that executes the correct interpreter sequence for the desired report
            match int(user_report_choice):
                # test case
                case 1:
                    self._valid = True
                    interpreter.test(report_name, interface_instance)
                # TBD case
                case 2:
                    self._valid = True
                # default case that signifies an invalid choice
                case _:
                    print("Not a valid choice")
                    self._valid = False
            # end match case
        # end while