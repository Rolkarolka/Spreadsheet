from class_spreadsheet import Spreadsheet
from sys import argv, exit, setrecursionlimit
import re
import os
from keys import Keys

if os.name == "nt":
    os.system('color 70')

else:
    print("\033[107m \033[30m")

setrecursionlimit(100)

class WrongNumberOfFiles(Exception):
    pass

def spreadsheet_help():
    print("TOPIC\nSpreadsheet\n")
    print("SHORT DESCRIPTION\nProgram which allows you to organize data in a spreadsheet.\n")
    print("OPTIONS\n")
    print(":> display \nPrint the table with data.\n")
    print(":> show <cell_address>\nShow  data in the selected cell_adress.\nexample:\n\t:>show A1\n\toutput: A1: [4]\n")
    print(":> set <cell_address> <cell_value>\nSet data in the selected cell_adress.\nexample:\n\t:>set A1 5\n\toutput: show table with new data.\n")
    print(":> edit <cell_address>\nPrint existing data and ask about new value.\nexample:\n\t:>edit A1\n\tA1: [5]\n\tNew value:_\n\toutput: show table with new data.\n")
    print(":> del <cell_address>\nDeletes the value assigned to the cell address.\nexample:\n\t:>del A4\n\toutput:delete values\n")
    print(":> size\nChanges the number of columns and rows.\nexample:\n\t:>size\n\toutput: show table with new size\n")
    print(":> add_row <quantity>\nAdds the specified number of rows.\nexample:\n\t:>add_row 11\n\toutput: show table with new rows\n")
    print(":> add_column <quantity>\nAdds the specified number of column.\nexample:\n\t:>add_column 4\n\toutput: show table with new column\n")
    print(":> move \nUse the arrows on the keyboard to move rows and columns.\nexample:\n\t:>move\n")
    print(":> cell_size <columns width>\n:> cell_size <column letter> <column width>\nSets the width of 1 columns or of all columns.\nexample:\n\t:>cell_size A 10\n\toutput: column A has now a width of 10\n\t:>cell_size 5\n\toutput: all column have now a width of 5\n")
    print(":> save <<optional_file_name>>\nSaves the data in the spreadsheet.\nexample:\n\t:>save\n\toutput: save your file\n")
    print(":> exit\nEnds the program.\n")
    print("AVAILABLE OPERATIONS ON NUMBERS:\n")
    print(" - addition\n\teg.: =5+3\n\t     =A1+3\n\t     =H11+AAA3")
    print(" - substraction\n\teg.: =5-4\n\t     =A1-5\n\t     =G12-S2")
    print(" - division\n\teg.: =3/2\n\t     =A1/4\n\t     =AZ3/Q1 ")
    print(" - multiplication\n\teg.: =3*2\n\t     =A1*5\n\t     =A4*A6")
    print(" - sum\n\teg.: =SUM(A1:A8)")
    print(" - average\n\teg.: =AVG(A1:C3)")
    print(" - minimum\n\teg.: =MIN(F1:F5)")
    print(" - maximum\n\teg.: =MAX(D2:D3)")


def get_command_list(input_data):
    command_list = input_data.split()
    while command_list == []:
        input_data = input("COMMAND:> ")
        command_list = input_data.split()
    else:
        return command_list


def check_display_params(params_list):
    if len(params_list) == 1:
        return True
    else:
        print("'display' doesn't require additional arguments")
        return False


def check_help_params(params_list):
    if len(params_list) == 1:
        return True
    else:
        print("'help' doesn't require additional arguments")
        return False


def check_save_params(param_list):
    if len(param_list) == 2:
        return True
    elif len(param_list) == 1:
        return True
    else:
        print(f"Wrong formula: {''.join(param_list)}. Pattern: save <<optional_file_name>> ")
        return False


def check_file_name(name):
    if re.match(r"(^[a-zA-Z0-9-_]+\.[a-z]+$)", os.path.basename(name)):
        return True
    else:
        return False


def get_file_name_from_input():
    file_name = input("Enter json file name: ")
    while not check_file_name(file_name):
        print("Wrong file name")
        file_name = input("Enter json file name: ")
    return file_name


def save_to_file(file_name=""):
    if file_name == "" and spreadsheet.file_name == "":
        file_name = get_file_name_from_input()
        spreadsheet.file_name = file_name
    elif file_name != "":
        if not check_file_name(file_name):
            file_name = get_file_name_from_input()
        spreadsheet.file_name = file_name
    spreadsheet.save_file()


def check_size_params(params_list):
    if len(params_list) == 1:
        return True
    else:
        print("'size' doesn't require additional arguments")
        return False


def check_add_row_params(params_list):
    if len(params_list) == 2 and params_list[1].isnumeric():
        return True
    else:
        print(f"Wrong formula: {' '.join(params_list)}. Pattern: add_row <quantity>")
        return False


def check_add_column_params(params_list):
    if len(params_list) == 2 and params_list[1].isnumeric():
        return True
    else:
        print(f"Wrong formula: {' '.join(params_list)}. Pattern: add_column <quantity>")
        return False


def check_move_params(params_list):
    if len(params_list) == 1:
        return True
    else:
        print("'size' doesn't require additional arguments")
        return False


def check_edit_params(params_list):
    if len(params_list) == 2 and re.match(r'(^[A-Z]+[1-9][0-9]*$)', params_list[1]):
        return True
    else:
        print(f"Wrong formula: {' '.join(params_list)}. Pattern: edit <cell_address>")
        return False


def check_cell_size_params(params_list):
    if len(params_list) == 3 and re.match(r'(^[a-zA-Z]+$)', params_list[1]) and params_list[2].isnumeric():
        return True
    elif len(params_list) == 2 and params_list[1].isnumeric():
        return True
    else:
        print(f"Wrong formula: {' '.join(params_list)}. Pattern 1: cell_size <columns width> 2: cell_size <column letter> <column width>")
        return False


def check_del_params(params_list):
    if len(params_list) == 2 and re.match(r'(^[A-Z]+[1-9][0-9]*$)', params_list[1]):
        return True
    else:
        print(f"Wrong formula: {' '.join(params_list)}. Pattern: del <cell_address>")
        return False


def check_set_params(params_list):
    if len(params_list) >= 3 and re.match(r'(^[A-Z]+[1-9][0-9]*$)', params_list[1]):
        return True
    else:
        print(f"Wrong formula: {' '.join(params_list)}. Pattern: set <cell_address> <cell_value>")
        return False


def check_show_params(params_list):
    if len(params_list) == 2 and re.match(r'(^[A-Z]+[1-9][0-9]*$)', params_list[1]):
        return True
    else:
        print(f"Wrong formula: {' '.join(params_list)}. Pattern: show <cell_address>")
        return False


def get_max_col():
    column = input("Enter the number of column: ")
    match_obj = re.match(r'(^[1-9][0-9]*$)', column)
    while not match_obj:
        print("Wrong value! I need integer larger than zero!")
        column = input("Enter the number of column: ")
        match_obj = re.match(r'(^[1-9][0-9]*$)', column)
    else:
        return int(column)


def get_max_row():
    row = input("Enter the number of rows: ")
    match_obj = re.match(r'(^[1-9][0-9]*$)', row)
    while not match_obj:
        print("Wrong value! I need integer larger than zero!")
        row = input("Enter the number of rows: ")
        match_obj = re.match(r'(^[1-9][0-9]*$)', row)
    else:
        return int(row)

if __name__ == "__main__":
    try:
        spreadsheet = Spreadsheet()
        key = Keys()
        spreadsheet.set_terminal_width(os.get_terminal_size().columns)
        spreadsheet.set_terminal_length(os.get_terminal_size().lines)
        if len(argv) == 2:
            if not spreadsheet.load(argv[1]):
                exit()
        elif len(argv) == 1:
            spreadsheet.set_max_col(get_max_col())
            spreadsheet.set_max_row(get_max_row())
        else:
            raise WrongNumberOfFiles("The wrong number of arguments was given")
        spreadsheet.display()

        spreadsheet._table_separator()
        command = get_command_list(input("COMMAND:>(help) "))

        while command[0] != "exit":
            if spreadsheet.get_terminal_width != os.get_terminal_size().columns or spreadsheet.get_terminal_length != os.get_terminal_size().linesa:
                spreadsheet.set_terminal_width(os.get_terminal_size().columns)
                spreadsheet.set_terminal_length(os.get_terminal_size().lines)
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            if command[0] == "show" and check_show_params(command):
                spreadsheet.cell_show(command[1])

            elif command[0] == "set" and check_set_params(command):
                word = " ".join(command[2:])
                spreadsheet.cell_set(command[1], word)
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            elif command[0] == "display" and check_display_params(command):
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            elif command[0] == "move" and check_move_params(command):
                if spreadsheet.max_col < spreadsheet.number_printed_column() and spreadsheet.max_row < spreadsheet.number_printed_rows():
                    print("The whole table is displayed on the screen")
                else:
                    print("If you want to finish, press enter")
                    if key.ended_key == []:
                        key.set_ended_key()
                    while key.new_key not in key.ended_key:
                        num = key.listen()
                        if (spreadsheet.max_col - spreadsheet.number_printed_column() + 1) > 0:
                            if (spreadsheet.max_col - spreadsheet.number_printed_column() + 1) <= num[0]:
                                num[0] -= 1
                            elif num[0] < 0:
                                num[0] = 0
                        else:
                            num[0] = spreadsheet.first_printed_column

                        if (spreadsheet.max_row - spreadsheet.number_printed_rows() + 1) > 0:
                            if (spreadsheet.max_row - spreadsheet.number_printed_rows() + 1) <= num[1]:
                                num[1] -= 1
                            elif num[1] < 0:
                                num[1] = 0
                        else:
                            num[1] = spreadsheet.first_printed_row
                        if num[0] != spreadsheet.first_printed_column:
                            spreadsheet.add_to_printed_first_column(num[0])
                            spreadsheet.add_to_printed_first_row(num[1])
                            os.system('cls' if os.name == 'nt' else 'clear')
                            spreadsheet.display()
                            print("If you want to finish, press enter")

                        if num[1] != spreadsheet.first_printed_row:
                            spreadsheet.add_to_printed_first_row(num[1])
                            os.system('cls' if os.name == 'nt' else 'clear')
                            spreadsheet.display()
                            print("If you want to finish, press enter")

                    key.new_key = None

            elif command[0] == "help" and check_help_params(command):
                spreadsheet_help()

            elif command[0] == "edit" and check_edit_params(command):
                spreadsheet.cell_edit(command[1])
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            elif command[0] == "del" and check_del_params(command):
                spreadsheet.cell_del(command[1])
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            elif command[0] == "size" and check_size_params(command):
                max_col = get_max_col()
                max_row = get_max_row()
                spreadsheet.set_spreadsheet_size(max_col, max_row)
                if max_col == spreadsheet.max_col and max_row == spreadsheet.max_row:
                    os.system('cls' if os.name == 'nt' else 'clear')
                    spreadsheet.display()

            elif command[0] == "add_column" and check_add_column_params(command):
                spreadsheet.add_to_column(command[1])
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            elif command[0] == "add_row" and check_add_row_params(command):
                spreadsheet.add_to_rows(command[1])
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            elif command[0] == "cell_size" and check_cell_size_params(command):
                if len(command) == 3:
                    spreadsheet.cell_size(command[1].upper(), command[2])
                elif len(command) == 2:
                    spreadsheet.set_default_width(command[1])
                os.system('cls' if os.name == 'nt' else 'clear')
                spreadsheet.display()

            elif command[0] == "save" and check_save_params(command):
                if len(command) == 2:
                    save_to_file(command[1])
                elif len(command) == 1:
                    save_to_file()
                print("Saved successfully")

            spreadsheet._table_separator()
            command = get_command_list(input("COMMAND:>"))

        save = input("Do you want to save file? (yes/no): ")
        while not save == "yes" and not save == "no":
            save = input("Do you want to save file? (yes/no): ")

        if save == "yes":
            save_to_file(spreadsheet.file_name)
        exit()
    except KeyboardInterrupt:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("Program was interrupted by user CTRL+C...")
