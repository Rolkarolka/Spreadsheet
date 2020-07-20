import json
import re

class WrongFileFormat(Exception):
    pass

class Spreadsheet:
    def __init__(self):
        self.max_col = 0
        self.max_row = 0
        self.default_width = 10
        self.file_name = ""
        self.content = {}
        self.individual_col_width = {}
        self.len_table_row = 0
        self._terminal_width = 123
        self._terminal_length = 123
        self.first_printed_column = 0
        self.first_printed_row = 0

    def number_printed_column(self):
        """ count how many column we can display """
        number_column = (self._terminal_width - 5 - len(str(self.max_col))) // (self.default_width + 1)
        # check up to date termianl width, substract 5 (intended for 2 blank space and 3 "|" in index_number) and divide by default_width + 1 ( intended for "|")
        return number_column

    def number_printed_rows(self):
        """ count how many rows we can display """
        number_rows = (self._terminal_length - 6) // 2
        # check up to date termianl length, substract 6 (intended for command_line, abc_index and separator) and divide by 2 because part of length is for separator
        return number_rows

    def add_to_printed_first_column(self, number: int):
        """ set the number of column from which we want to start displaying """
        if isinstance(number, int):
            if (self.max_col - number + 1) >= self.number_printed_column():   # +1 <- first column for index_number
                self.first_printed_column = number
        else:
            raise TypeError("The number of column must be integer")

    def add_to_printed_first_row(self, number: int):
        """ set number of row from which we want to start displaying """
        if isinstance(number, int):
            if (self.max_row - number + 1) >= self.number_printed_rows():   # +1 <- first column for index_abc
                self.first_printed_row = number
        else:
            raise TypeError("The number of row must be integer")

    def set_max_col(self, max_col):
        """ set number of column """
        if isinstance(max_col, int):
            self.max_col = max_col
        else:
            raise TypeError("The number of column must be integer")

    def set_max_row(self, max_row):
        """ set number of rows """
        if isinstance(max_row, int):
            self.max_row = max_row
        else:
            raise TypeError("The number of column must be integer")

    def get_terminal_width(self):
        return self._terminal_width

    def get_terminal_length(self):
        return self._terminal_length

    def set_terminal_width(self, terminal_width):
        if isinstance(terminal_width, int) or isinstance(terminal_width, float):
            self._terminal_width = terminal_width - 1   # for  correct display (1 sign for "\n")
        else:
            raise TypeError("Terminal_width must be integer or float")

    def set_terminal_length(self, terminal_length):
        if isinstance(terminal_length, int) or isinstance(terminal_length, float):
            self._terminal_length = terminal_length
        else:
            raise TypeError("Terminal length must be integer or float")

    def _is_cell(self, cell: str):
        """ check if inupt data is cell e.g. A2 """
        match_obj = re.match(r"(^[A-Z]+[1-9][0-9]*$)", str(cell))
        if match_obj:
            return True
        else:
            return False

    def _is_int(self, string):
        """ check if in string is integer"""
        match_obj = re.match(r"(^[0-9]+$)", str(string))
        if match_obj:
            return True
        else:
            return False

    def _num_to_letter(self, number: int, letter=""):
        # A = 1
        # Z = 26
        # AA = 27 ...

        if number < 1:   # offset so that the letter "A" takes the value 1
            return " "
        elif 1 <= number < 27:
            letter += (chr(64 + number))  # in ASCII, the alphabet assumes values ​​65 - 90
        elif number >= 27:   # if number >= 27  function divides into two appropriate numbers and creates string from received letters
            number -= 1     # related to offset
            second = 1 + number % 26
            first = (number // 26)
            letter = self._num_to_letter(first, letter)
            letter = self._num_to_letter(second, letter)
        return letter

    def _letter_to_num(self, letter: str, number=0):
        # 1 = A
        # 26 = Z
        # 27 == AA ...
        if len(letter) == 1:
            number += ord(letter) - 64   # offset so that the letter "A" takes the value 1
        else:
            number += self._letter_to_num(letter[-1])   # divides string into letters and converts them according to ASCII into a number
            number += self._letter_to_num(letter[0:-1]) * 26
        return number

    def load(self, document):
        """ loads a file that must contain max_col and max_row, may also contain default_width,
        col_width - dictionary which contain idividual column width and data"""
        try:
            with open(str(document), "r") as file:
                try:
                    tmp_content = json.load(file)
                except ValueError:
                    return False
                if ("max_col" in tmp_content) and ("max_row" in tmp_content) and ("default_width" in tmp_content) and ("col_width" in tmp_content) and "data" in tmp_content:
                    if isinstance(tmp_content["max_col"], int) and isinstance(tmp_content["max_row"], int) and isinstance(tmp_content["default_width"], int):
                        self.max_col = tmp_content["max_col"]
                        self.max_row = tmp_content["max_row"]
                        self.default_width = tmp_content["default_width"]
                        self.individual_col_width = tmp_content["col_width"]
                        self.content = tmp_content["data"]
                        self.file_name = document
                        return True
                    else:
                        raise TypeError("The number of column and rows must be integer")
                else:
                    raise WrongFileFormat("File must contain max_col and max_row")
        except FileNotFoundError as fex:
            print(f"Error! {fex}.")
            return False

    def save_file(self):
        """ collects all informations about your spreasheet and save to json file"""
        current_content = {}
        current_content["max_row"] = self.max_row
        current_content["max_col"] = self.max_col
        current_content["default_width"] = self.default_width
        current_content["data"] = self.content
        current_content["col_width"] = self.individual_col_width
        try:
            with open(self.file_name, "w") as outfile:
                json.dump(current_content, outfile, sort_keys=True, indent=4)
        except FileNotFoundError:
            print(f"File name cannot be {self.file_name}")

    def get_numeric_value_from_string(self, number: str):
        try:
            if "." in number:
                return float(number)
            else:
                return int(number)
        except TypeError as tx:
            print(f"Error! {tx}")

    def get_cell_value(self, param):
        match_obj = re.match(r"^([A-Z]+[0-9]+)$", param)
        if match_obj:
            if param in self.content:
                if param != self.content[param]:
                    return self.get_real_cell_value(str(self.content[param]))
                else:
                    return "Err rec"
            else:
                return ""
        else:
            return self.get_numeric_value_from_string(param)

    def get_range_to_list(self, pattern, string):
        """ function which change cell address into the vector [x, y] """
        match_obj = re.match(pattern, string)
        if match_obj:
            x_axis = [self._letter_to_num(match_obj.group("left_letter")), self._letter_to_num(match_obj.group("right_letter"))]
            y_axis = [int(match_obj.group("left_number")), int(match_obj.group("right_number"))]
            x_axis.sort()
            y_axis.sort()
            numbers = []
            for x in range(x_axis[0], x_axis[1] + 1):  # +1 so that it also contains the last number
                for y in range(y_axis[0], y_axis[1] + 1):
                    cell_address = self._num_to_letter(x) + str(y)
                    if cell_address in self.content:
                        value = self.get_real_cell_value(cell_address)
                        if isinstance(value, str):
                            value = self.get_real_cell_value(value)
                        if isinstance(value, int) or isinstance(value, float):
                            numbers.append(value)
            return numbers
        else:
            return None

    def get_real_cell_value(self, cell_content):
        try:
            """ function which get real cell value if entered formula is correct"""
            # check if in cell is "plain number" i.e. (3), (=5), (=7.4), (=-2.5) etc.
            match_obj = re.match(r"^=?(?P<number>[-+]?[0-9]*\.?[0-9]+)$", cell_content)
            if match_obj:
                return self.get_numeric_value_from_string(match_obj.group("number"))

            # check if in cell is "single cell address" (=A1), (=F5) etc.
            match_obj = re.match(r"^=?(?P<address>([A-Z]+[0-9]+))$", cell_content)
            if match_obj:
                if match_obj.group("address") in self.content:
                    return str(self.get_real_cell_value(str(self.content[match_obj.group("address")])))
                else:
                    return ""

            # check if in cell is "equation" with numbers or cell addresses i.e. (=3+2), (=5*6), (=7.4/2), (=-2.5-5), (=A1-5), (=B6*5.2), (=-2.5+C1), (=A3-B4) etc.
            match_obj = re.match(r"^=(?P<left>([A-Z]+[0-9]+)|[-+]?[0-9]*\.?[0-9]+)(?P<operator>([\+|\-|\*|\/]))(?P<right>([A-Z]+[0-9]+)|[-+]?[0-9]*\.?[0-9]+)$", cell_content)
            if match_obj:
                operator = match_obj.group("operator")
                left_param = self.get_cell_value(match_obj.group("left"))
                right_param = self.get_cell_value(match_obj.group("right"))

                if left_param == "":
                    left_param = 0
                elif isinstance(left_param, str):
                    left_param = self.get_real_cell_value(left_param)
                if right_param == "":
                    right_param = 0
                elif isinstance(right_param, str):
                    right_param = self.get_real_cell_value(right_param)
                try:
                    if operator == "+":
                        return left_param + right_param
                    elif operator == "-":
                        return left_param - right_param
                    elif operator == "*":
                        return left_param * right_param
                    elif operator == "/":
                        if right_param != 0:
                            return round((left_param / right_param), 3)
                        else:
                            return "Err DIV 0"
                    else:
                        return "ERROR"
                except TypeError:
                    return f"Err num{operator}str"

            # calculate SUM i.e. (=SUM(A1:B3))
            numbers_list = self.get_range_to_list(r"^=SUM\((?P<left_letter>([A-Z]+))(?P<left_number>([0-9]+)):(?P<right_letter>([A-Z]+))(?P<right_number>([0-9]+))\)$", cell_content)
            if numbers_list is not None:
                sum = 0
                for number in numbers_list:
                    sum += number
                return sum

            # calculate AVG (=AVG(A1:B3))
            numbers_list = self.get_range_to_list(r"^=AVG\((?P<left_letter>([A-Z]+))(?P<left_number>([0-9]+)):(?P<right_letter>([A-Z]+))(?P<right_number>([0-9]+))\)$", cell_content)
            if numbers_list is not None:
                sum = 0
                for number in numbers_list:
                    sum += number
                return round((sum) / (len(numbers_list)), 3)

            # calculate MIN (=MIN(A1:A3))
            numbers_list = self.get_range_to_list(r"^=MIN\((?P<left_letter>([A-Z]+))(?P<left_number>([0-9]+)):(?P<right_letter>([A-Z]+))(?P<right_number>([0-9]+))\)$", cell_content)
            if numbers_list is not None:
                min_number = min(numbers_list)
                return min_number

            # calculate MAX (=MAX(A1:A3))
            numbers_list = self.get_range_to_list(r"^=MAX\((?P<left_letter>([A-Z]+))(?P<left_number>([0-9]+)):(?P<right_letter>([A-Z]+))(?P<right_number>([0-9]+))\)$", cell_content)
            if numbers_list is not None:
                max_number = max(numbers_list)
                return max_number

            # if any above match - return as is
            return cell_content
        except RecursionError:
            return f'ERR Rec'

    def _table_separator(self):
        """ make separator which looks like '----' """
        width = 0
        len_individual_width = 0
        for individual_width in self.individual_col_width.keys():
            width += self.individual_col_width[individual_width]
            len_individual_width += 1  # because every cell is separated by "|"
        size_table_column = 5 + len(str(self.max_row)) + width + len_individual_width + ((self.default_width + 1) * (self.max_col - len_individual_width - self.first_printed_column))
        if size_table_column < self._terminal_width:
            print("-" * (size_table_column))
        else:
            print("-" * (self._terminal_width))

    def _abc_index(self):
        """ print index of letters """
        letter_column = ""
        for x in range(self.first_printed_column, self.max_col + 1):
            if x == self.first_printed_column:
                print("")
                letter_column = "|" + self._num_to_letter(x - self.first_printed_column).center(2 + len(str(self.max_row))) + "||"
            else:
                x = self._num_to_letter(x)
                if x in self.individual_col_width.keys():
                    width = self.individual_col_width[x]
                else:
                    width = self.default_width
                letter_column += x.center(width) + "|"
        if len(letter_column) < self._terminal_width:
            print(letter_column)
        else:
            print(letter_column[:self._terminal_width])
        self._table_separator()
        return letter_column

    def _rest_table(self):
        """ print rest information for your table """
        if self.number_printed_rows() > self.max_row:
            table_length = self.max_row
        else:
            table_length = self.first_printed_row + self.number_printed_rows()

        if self.number_printed_column() > self.max_col:
            table_width = self.max_col
        else:
            table_width = self.first_printed_column + self.number_printed_column()

        for y in range(self.first_printed_row + 1, table_length + 1):
            table = ""
            for x in range(self.first_printed_column, table_width + 1):
                if x == self.first_printed_column:
                    table = table + "|" + str(y).center(2 + len(str(self.max_row))) + "||"
                else:
                    cell_address = self._num_to_letter(x) + str(y)

                    if cell_address[0] in self.individual_col_width.keys():
                        width = self.individual_col_width[cell_address[0]]
                    else:
                        width = self.default_width
                    if cell_address in self.content:
                        real_cell_value = str(self.get_real_cell_value(str(self.content[cell_address])))
                        if len(real_cell_value) > width:
                            table = table + "####".center(width) + "|"
                        else:
                            table = table + real_cell_value.center(width) + "|"
                    else:
                        table = table + "".center(width) + "|"
            self.len_table_row = len(table)
            print(table[:self._terminal_width], end="")
            print("")
            self._table_separator()

    def set_spreadsheet_size(self, max_col, max_row):
        """ - function for user """
        if isinstance(max_col, int) and isinstance(max_row, int):
            if (max_col >= self.max_col) and (max_row >= self.max_row):
                self.set_max_col(max_col)
                self.set_max_row(max_row)
            else:
                print("New max_row and max_column must be larger than before.")
        else:
            print("Number of column or rows must be integer")

    def display(self):
        """ display table - function for user"""
        if self.max_col != 0 and self.max_row != 0:
            self._abc_index()
            self._rest_table()
        else:
            print("Number of column or rows cannot equal zero")

    def cell_show(self, cell):
        """ show formula and real value of cell - function for user """
        if self._is_cell(cell):
            if cell in self.content:
                print(cell + ": [" + str(self.content[cell]) + "] (" + str(self.get_real_cell_value(str(self.content[cell]))) + ")")
            else:
                print(cell + ": Empty")
        else:
            print("No matching format")

    def cell_set(self, cell, value):
        """ set value of cell - function for user """
        if self._is_cell(cell):
            self.content[cell] = value
        else:
            print("No matching format")

    def cell_edit(self, cell):
        """ edit value of cell - function for user """
        if self._is_cell(cell):
            self.cell_show(cell)
            new_value = input("New value: ")
            self.cell_set(cell, new_value)
        else:
            print("No matching format")

    def cell_del(self, cell):
        """ del value of cell - function for user recieved from input"""
        if self._is_cell(cell) and (cell in self.content):
            del self.content[cell]
        elif not self._is_cell(cell):
            print("No matching format")
        else:
            print("Cell is already empty")

    def cell_size(self, column, number):
        """ set size of chosen column - function for user recieved from input """
        if self._is_int(column):
            column = self._num_to_letter(int(column))
        for col_number in range(0, self.max_col):
            if column == self._num_to_letter(col_number):
                self.individual_col_width[column] = int(number)
        else:
            print(f"Column: {column} is not in the range (1, {self.max_col} = {self._num_to_letter(self.max_col)})")

    def set_default_width(self, number):
        """ set width of all cells - function for user recieved from input """
        if self._is_int(number) and len(self._num_to_letter(self.max_col)) < int(number):
            self.default_width = int(number)
        else:
            print("Default width must be integer")

    def add_to_column(self, added_column):
        """ increase number of column - function for user recieved from input """
        if self._is_int(added_column):
            self.max_col += int(added_column)
        else:
            print("The number of added column must be integer.")

    def add_to_rows(self, added_rows):
        """ increase number of rows - function for user recieved from input """
        if self._is_int(added_rows):
            self.max_row += int(added_rows)
        else:
            print("The number of added column must be integer.")
