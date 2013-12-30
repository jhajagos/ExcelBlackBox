__author__ = 'janos'


"""
This program is for evaluating Excel spreadsheets and capturing the input from
sometimes crazy formulas in the spreadsheet.

An Excel Spreadsheet can be considered as a multivariate function

Function(A1, A2, . . .) = (B1, B2, . .)

The program is designed to capture the relationship between input and outputs.

As an example, from
A1, A2, B1
so to allow further analysis of the input.
"""


try:
    import win32
except ImportError:
    pass

from math import sin


class ExcelDriver(object):
    """Parent class for the Microsoft Excel Driver for Python"""

    def open_workbook(self, file_name):
        pass

    def set_cell_value(self, cell_address, value_to_set, worksheet_name=None):
        self._set_cell_value(cell_address, value_to_set, worksheet_name)

    def read_cell_value(self, cell_address, worksheet_name=None):
        return self._read_cell_value(cell_address, worksheet_name)


class ExcelDriverCOM(ExcelDriver):
    """Driver that utilizes the COM interface in Microsoft Windows to script Excel"""

    def __init__(self):
        self.excel = win32.gencache.EnsureDispatch("Excel.Application")
        self.worksheets = {}

    def open_workbook(self, file_name):
        self.workbook = self.excel.Workbooks.open(file_name)

    def set_worksheet(self, worksheet_name="Sheet1"):
        self.current_worksheet_name = worksheet_name
        self.worksheet = self.workbook.Worksheets(worksheet_name)

        self.worksheets[worksheet_name] = self.worksheet

    def _set_cell_value(self, cell_address, value_to_set, worksheet_name):

        if worksheet_name != self.current_worksheet_name:
            self.set_worksheet(worksheet_name)

        self.worksheet.Range(cell_address).Value = value_to_set

    def _read_cell_value(self, cell_address, worksheet_name):

        if worksheet_name != self.current_worksheet_name:
            self.set_worksheet(worksheet_name)

        return self.worksheet.Range(cell_address).Value


def example_excel_dummy_function(cell_address_dict):
    try:
        a = cell_address_dict["B2"]
        b = cell_address_dict["B3"]
        theta = cell_address_dict["B4"]
    except KeyError:
        return None

    return a * sin((1/b) * theta)


class ExcelDriverDummy(ExcelDriver):
    """The Excel Dummy Driver is for testing purposes"""
    def __init__(self, excel_function=example_excel_dummy_function, excel_cell_value_to_update="C2"):
        self.excel_dict = {}
        self.excel_function = excel_function
        self.excel_cell_value_to_update = excel_cell_value_to_update

    def _update_model(self):
        self.excel_dict[self.excel_cell_value_to_update] = self.excel_function(self.excel_dict)

    def open_workbook(self, file_name):
        pass

    def set_worksheet(self, worksheet_name="Sheet1"):
        pass

    def _set_cell_value(self, cell_address, value_to_set, worksheet):
        self.excel_dict[cell_address] = value_to_set

    def _read_cell_value(self, cell_address, worksheet_name):
        self._update_model()
        return self.excel_dict[cell_address]


class ExcelBlackBoxEvaluator(object):
    """An evaluator for a struct"""

    def __init__(self, input_variable_dict, output_variable_dict, excel_workbook_file_name, excel_driver=ExcelDriverDummy):
        """
    input_variable_dict = { "a" : ("a2", sheet1) }
    output_variable_dict = { "fsin" : ("b2","sheet1")}
    excel_workbook_filename = "file_name.xlsx"
        """
        self.excel_workbook_file_name = excel_workbook_file_name
        self.input_variable_dict = input_variable_dict
        self.output_variable_dict = output_variable_dict
        self.excel_driver = excel_driver()
        self.excel_driver.open_workbook(self.excel_workbook_file_name)

    def evaluate(self, input_parameters):
        """{"a": 5.0"}"""
        for input_parameter_key in input_parameters:
            input_parameter_value = input_parameters[input_parameter_key]
            input_parameter_address = self.input_variable_dict[input_parameter_key]
            self.excel_driver.set_cell_value(input_parameter_address[0], input_parameter_value, input_parameter_address[1])

        result_dict = {}
        for output_variable in self.output_variable_dict:
            output_value_address, output_value_worksheet = self.output_variable_dict[output_variable]
            result_dict[output_variable] = \
                self.excel_driver.read_cell_value(output_value_address, output_value_worksheet)

        return result_dict