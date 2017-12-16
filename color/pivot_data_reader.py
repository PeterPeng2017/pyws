import logging

import xlrd

from compare.app_constants import AppConstants

logger = logging.getLogger( 'SalesDataReader')
logger.setLevel(logging.DEBUG)


class PivotDataContainer:
    def __init__(self, sheet_data):
        self.sheet_data = sheet_data


class CellUnit:
    def __init__(self, value, x, y, color):
        self.value = value
        self.x = x
        self.y = y
        self.color = color


class PivotDataReader:

    def __init__(self, file_path, first_row, max_row, max_col, skip_cols):
        self.file_path = file_path
        self.first_row = first_row
        self.max_row = max_row
        self.max_col = max_col
        self.skip_cols = skip_cols
        self.excel_book = xlrd.open_workbook(file_path)
        self.duplicate_keys = []

    def read_sheet_data(self, sheet_name):
        logger.info("read file: {0}".format(self.file_path))
        sheet = self.excel_book.sheet_by_name(sheet_name)

        sheet_data = {}

        account_id = "NONE"
        for row_index in range(self.first_row, self.max_row):
            account_data = sheet.cell(row_index, 0).value
            if len(str(account_data)) > 0:
                account_id = account_data

            row_key = account_id + AppConstants.KEY_SEPARATOR + str(sheet.cell(row_index, 2).value)

            if row_key in sheet_data:
                logger.error("duplicate key:" + row_key)
                self.duplicate_keys.append(row_key)

            row_data = sheet_data.setdefault(row_key, {})
            for col_index in range(4, self.max_col):
                if col_index not in self.skip_cols:
                    cell_unit = CellUnit(sheet.cell(row_index, col_index).value,
                                         row_index + 1,
                                         col_index + 1,
                                         AppConstants.NONE)
                    row_data[col_index] = cell_unit

        pivot_data_container = PivotDataContainer(sheet_data)
        return pivot_data_container


