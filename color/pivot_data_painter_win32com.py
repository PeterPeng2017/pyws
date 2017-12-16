from win32com.client import Dispatch
from compare.app_constants import AppConstants
import win32com.client


class PivotDataPainterWin32com:

    def __init__(self,
                 crm_file_path,
                 crm_sheet_name,
                 crm_result_file_path,
                 sap_file_path,
                 sap_sheet_name,
                 sap_result_file_path,
                 crm_data_container,
                 sap_data_container):

        self.crm_file_path = crm_file_path
        self.crm_sheet_name = crm_sheet_name
        self.crm_result_file_path = crm_result_file_path
        self.sap_file_path = sap_file_path
        self.sap_sheet_name = sap_sheet_name
        self.sap_result_file_path = sap_result_file_path
        self.crm_data_container = crm_data_container
        self.sap_data_container = sap_data_container

    def save(self):
        self.paint_workbook(self.crm_file_path,
                            self.crm_sheet_name,
                            self.crm_result_file_path,
                            self.crm_data_container)

        self.paint_workbook(self.sap_file_path,
                            self.sap_sheet_name,
                            self.sap_result_file_path,
                            self.sap_data_container)


    @staticmethod
    def paint_workbook(file_path, sheet_name, result_file_path, data_container):

        xlApp = win32com.client.DispatchEx("Excel.Application")

        xlBook = xlApp.Workbooks.Open(file_path)

        sheet = xlBook.Worksheets(sheet_name)

        sheet_data = data_container.sheet_data
        for row_key, row_data in sheet_data.items():
            for cell_key, cell_unit in row_data.items():
                if cell_unit.color != AppConstants.NONE:
                    color_win32 = PivotDataPainterWin32com.convert_color_to_win32(cell_unit.color)
                    sheet.Cells(cell_unit.x, cell_unit.y).Interior.Color = color_win32

        xlBook.SaveAs(result_file_path)

        xlBook.Close(SaveChanges=0)
        xlApp.Application.Quit()

    #Attention that  color is GBR not RGB.
    @staticmethod
    def convert_color_to_win32(color_str):
        rgb = (255, 255, 255)
        if color_str == AppConstants.RED:
            rgb = (0, 0, 255)
        elif color_str == AppConstants.BLUE:
            rgb = (195, 135, 1)
        elif color_str == AppConstants.GREEN:
            rgb = (2, 139, 1)
        elif color_str == AppConstants.YELLOW:
            rgb = (2, 232, 253)

        return PivotDataPainterWin32com.rgb_to_hex(rgb)

    @staticmethod
    def rgb_to_hex(rgb):
        strValue = '%02x%02x%02x' % rgb
        iValue = int(strValue, 16)
        return iValue
