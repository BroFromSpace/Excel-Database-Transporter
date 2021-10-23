import pandas as pd
import xlsxwriter
import os

# Define paths to your excel spreadsheets
from typing import List, Any, Tuple

DB_START_FILE = "database_start.xlsx"  # Name of start excel file
DB_RESULT_FILE = "database_result.xlsx"  # Name of result excel file
PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))  # Project root path


def get_part_list() -> List[Tuple[Any]]:
    """
    Create product list from start file

    :rtype: object
    :return: list of products
    """
    db_start_path = os.path.join(PROJECT_ROOT, DB_START_FILE)  # path to start file
    # load excel file with pandas
    xl = pd.ExcelFile(db_start_path)
    # get all sheets in excel file
    sheets = xl.sheet_names
    return [
        part
        for sheet in sheets
        for part in zip(
            *[xl.parse(sheet).iloc[:, i] for i in range((xl.parse(sheet)).shape[1])]
        )
    ]


def copy_part(part: Tuple[Any], part_number: str, part_count: int) -> bool:
    """
    Try to copy row from start database to result database
    :param part:
    :param part_number:
    :param part_count:
    :rtype: bool
    :return: true if row was copied successfully else false
    """

    db_result_path = os.path.join(PROJECT_ROOT, DB_RESULT_FILE)  # path to result file

    part = list(part)
    part.insert(0, part_number)

    try:
        # read data from excel file and make dataFrame from it
        dfs = pd.read_excel(db_result_path, sheet_name="Sheet1", engine="openpyxl")

        # create 'Part storage number' column if it is not exist
        if 'Part storage number' not in dfs:
            dfs.insert(0, "Part storage number", " ")
        ls = dfs.to_dict(orient='split')
        data = (ls['data'])

        for _ in range(part_count):
            data.append(part)

        # open excel result file and create new worksheet in it
        workbook = xlsxwriter.Workbook(db_result_path)
        worksheet = workbook.add_worksheet()
        # create columns
        worksheet.write_row(0, 0, ['Part storage number', 'Part', 'Quantity', 'Category'])
        # write all rows into new excel file
        for ct, i in enumerate(data, start=1):
            worksheet.write_row(ct, 0, i)

        # save and close result excel file
        workbook.close()
    except Exception as e:
        print(e)
        return False
    else:
        return True
