"""Package for handling Excel file"""
import openpyxl


class ReadHelper:
    """OpenpyxlHelper"""

    def __init__(self) -> None:
        pass

    @staticmethod
    def get_sheet_names_by_file_path(path: str) -> [str]:
        """
        path: file path
        """
        book = openpyxl.open(path)
        return book.sheetnames
