from win32com.client import Dispatch
from color.pivot_data_painter_win32com import PivotDataPainterWin32com
from compare.app_constants import AppConstants
import win32com.client



def rgb_to_hex(rgb):
    strValue = '%02x%02x%02x' % rgb
    iValue = int(strValue, 16)
    return iValue


xlApp = win32com.client.DispatchEx("Excel.Application")

xlBook = xlApp.Workbooks.Open("F:\\python\\compareCRMAndSAP\\color1\\CRM.xlsx")

sheet = xlBook.Worksheets("By Detailed Product")

sheet.Cells(13, 5).Interior.Color = PivotDataPainterWin32com.convert_color_to_win32(AppConstants.YELLOW)
sheet.Cells(13, 6).Interior.Color = PivotDataPainterWin32com.convert_color_to_win32(AppConstants.BLUE)
sheet.Cells(13, 9).Interior.Color = PivotDataPainterWin32com.convert_color_to_win32(AppConstants.GREEN)
sheet.Cells(13, 10).Interior.Color = PivotDataPainterWin32com.convert_color_to_win32(AppConstants.RED)

print(sheet.Cells(13, 5))

xlBook.SaveAs("F:\\python\\compareCRMAndSAP\\color1\\crm_test.xlsx")


xlBook.Close(SaveChanges=0)
xlApp.Application.Quit()


