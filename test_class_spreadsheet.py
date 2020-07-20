import pytest
from class_spreadsheet import Spreadsheet, WrongFileFormat

def test_cell_size():
    spreadsheet = Spreadsheet()
    spreadsheet.set_max_col(10)
    spreadsheet._abc_index()

    spreadsheet.cell_size("A", 4)
    assert spreadsheet.individual_col_width['A'] == 4

    spreadsheet.cell_size("4", 3)
    assert spreadsheet.individual_col_width["D"] == 3

    spreadsheet.cell_size(3, 3)
    assert spreadsheet.individual_col_width["C"] == 3

    spreadsheet.cell_size(4, "6")
    assert spreadsheet.individual_col_width["D"] == 6

def test_set_default_width():
    spreadsheet = Spreadsheet()
    spreadsheet.set_default_width(5)
    assert spreadsheet.default_width == 5

    spreadsheet.set_default_width("3")
    assert spreadsheet.default_width == 3


def test_add_to_printed_first_column():
    spreadsheet = Spreadsheet()
    spreadsheet.number_printed_column()

    spreadsheet.set_max_col(5)
    spreadsheet.add_to_printed_first_column(10)
    assert spreadsheet.first_printed_column == 0

    spreadsheet.set_max_col(30)
    spreadsheet.add_to_printed_first_column(4)
    assert spreadsheet.first_printed_column == 4

    with pytest.raises(TypeError):
        spreadsheet.add_to_printed_first_column(3.5)

def test_add_to_printed_first_row():
    spreadsheet = Spreadsheet()
    spreadsheet.number_printed_rows()

    spreadsheet.set_max_row(5)
    spreadsheet.add_to_printed_first_row(8)
    assert spreadsheet.first_printed_row == 0

    spreadsheet.set_max_row(100)
    spreadsheet.add_to_printed_first_row(10)
    assert spreadsheet.first_printed_row == 10
    with pytest.raises(TypeError):
        spreadsheet.add_to_printed_first_row(3.5)


def test_set_max_col():
    spreadsheet = Spreadsheet()
    spreadsheet.set_max_col(10)
    assert spreadsheet.max_col == 10

def test_set_max_row():
    spreadsheet = Spreadsheet()
    spreadsheet.set_max_row(15)
    assert spreadsheet.max_row == 15


def test_cell_set():
    spreadsheet = Spreadsheet()
    spreadsheet.cell_set("A1", 5)
    assert spreadsheet.content["A1"] == 5

def test_add_to_rows():
    spreadsheet = Spreadsheet()

    spreadsheet.add_to_rows("5")
    assert spreadsheet.max_row == 5

    spreadsheet.add_to_rows(5)
    assert spreadsheet.max_row == 10


def test_add_to_column():
    spreadsheet = Spreadsheet()

    spreadsheet.add_to_column("3")
    assert spreadsheet.max_col == 3

    spreadsheet.add_to_column(4)
    assert spreadsheet.max_col == 7

def test_cell_del():
    spreadsheet = Spreadsheet()

    spreadsheet.content["A1"] = 5
    spreadsheet.cell_del("A1")
    assert "A1" not in spreadsheet.content

def test_get_real_cell_value():
    spreadsheet = Spreadsheet()
    spreadsheet.content = {"A1": "4", "A2": 0, "A4": "Ania", "A3": "=J3", "A7": "=MIN(A1:A6)", "A8": "=SUM(A1:A3)", "A9": "=3+5", "B1": "=A1/2", "B2": "=A7*A1", "B3": "=AVG(A1:A7)", "B5": "=A1-100", "B6": "=MAX(A1:B3)"}
    assert spreadsheet.get_real_cell_value(spreadsheet.content["A1"]) == 4
    assert spreadsheet.get_real_cell_value(spreadsheet.content["A4"]) == "Ania"
    assert spreadsheet.get_real_cell_value(spreadsheet.content["A3"]) == ""
    assert spreadsheet.get_real_cell_value(spreadsheet.content["A7"]) == 0
    assert spreadsheet.get_real_cell_value(spreadsheet.content["A8"]) == 4
    assert spreadsheet.get_real_cell_value(spreadsheet.content["A9"]) == 8
    assert spreadsheet.get_real_cell_value(spreadsheet.content["B1"]) == 2
    assert spreadsheet.get_real_cell_value(spreadsheet.content["B3"]) == 1.333
    assert spreadsheet.get_real_cell_value(spreadsheet.content["B2"]) == 0
    assert spreadsheet.get_real_cell_value(spreadsheet.content["B5"]) == -96
    assert spreadsheet.get_real_cell_value(spreadsheet.content["B6"]) == 4

def test_load():
    spreadsheet = Spreadsheet()
    spreadsheet.load("test_file.json")
    assert spreadsheet.individual_col_width != {}
    assert spreadsheet.content != {}
    assert spreadsheet.default_width != 0
    assert spreadsheet.max_col != 0
    assert spreadsheet.max_row != 0
    assert spreadsheet.file_name == "test_file.json"
