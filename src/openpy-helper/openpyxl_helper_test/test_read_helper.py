"""unitest"""
import unittest

from openpyxl_helper.read_helper import ReadHelper


class TestReadHelper(unittest.TestCase):
    """TestOpenpyxlHelper"""

    def test_get_sheet_names_by_file_path(self):
        """test_get_sheet_names_by_file_path"""
        input_file_path = "../../test/file/ReadTest.xlsx"
        excepted_sheet_names = ["工作表1", "工作表2", "工作表3"]
        self.assertEqual(
            excepted_sheet_names,
            ReadHelper.get_sheet_names_by_file_path(input_file_path),
        )
