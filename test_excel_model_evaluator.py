__author__ = 'janos'

import unittest
import math

from excel_model_evaluator import ExcelBlackBoxEvaluator, ExcelDriverDummy, ExcelDriverCOM
import os


class TestExcelDriverDummy(unittest.TestCase):
    def setUp(self):
        self.excel_driver_obj = ExcelDriverDummy()

    def test_basic_read_write(self):
        self.excel_driver_obj.set_cell_value("A1", 1.0, "Sheet1")
        self.excel_driver_obj.set_cell_value("A2", 2.0, "Sheet1")
        self.excel_driver_obj.set_cell_value("A3", 2.0, "Sheet1")

        a = self.excel_driver_obj.read_cell_value("A1", "Sheet1")
        self.assertEquals(1.0, a)

        b = self.excel_driver_obj.read_cell_value("A2", "Sheet1")
        self.assertEquals(2.0, b)

    def test_function_evaluation(self):

        self.excel_driver_obj.set_cell_value("B1", 1.5, "Sheet1")
        self.excel_driver_obj.set_cell_value("B2", 2.0, "Sheet1")
        self.excel_driver_obj.set_cell_value("B3", 2.0, "Sheet1")

        result = self.excel_driver_obj.read_cell_value("C2")
        self.assertIsNone(result)
        self.excel_driver_obj.set_cell_value("B4", math.pi, "Sheet1")
        result2 = self.excel_driver_obj.read_cell_value("C2")
        self.assertIsNotNone(result2)

        expected_result2 = 2.0 * math.sin( (1/2.0) * math.pi)

        self.assertEquals(expected_result2, result2)


class TestExcelExcelBlackBoxEvaluator(unittest.TestCase):
    def setUp(self):
        pass

    def test_evaluate_dummy(self):
        sheet = "sheet1"
        input_parameters = {"a": ("B2", sheet), "b": ("B3", sheet), "theta": ("B4", sheet)}
        output_parameters = {"feval": ("C2", sheet)}
        workbook_name = "test_workbook.xlsx"

        excel_black_obj = ExcelBlackBoxEvaluator(input_parameters, output_parameters, workbook_name)

        to_evaluate = {"a": 2.0, "b": 2.0, "theta": math.pi}

        result = excel_black_obj.evaluate(to_evaluate)

        self.assertIsNotNone(result)
        expected_result = {"feval": 2.0 * math.sin((1/2.0) * math.pi)}

        self.assertEquals(expected_result, result)

    def test_evaluate(self):

        directory, script_name = os.path.split(os.path.abspath(__file__))
        workbook_name = os.path.join(directory, "test_model.xlsx")
        sheet = "Sheet1"
        input_parameters = {"a": ("B2", sheet), "b": ("B3", sheet), "theta": ("B4", sheet)}
        output_parameters = {"feval": ("C2", sheet)}

        excel_black_obj = ExcelBlackBoxEvaluator(input_parameters, output_parameters, workbook_name, ExcelDriverCOM)

        to_evaluate = {"a": 2.0, "b": 2.0, "theta": math.pi}

        result = excel_black_obj.evaluate(to_evaluate)

        self.assertIsNotNone(result)
        expected_result = {"feval": 2.0 * math.sin((1/2.0) * math.pi)}

        self.assertEquals(expected_result, result)

        excel_black_obj.close()



if __name__ == "__main__":
    unittest.main()