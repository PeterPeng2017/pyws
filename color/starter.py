import logging

from color.pivot_data_reader import PivotDataContainer
from color.pivot_data_reader import PivotDataReader
from color.pivot_data_comparator import PivotDataComparator
from color.pivot_data_painter_win32com import PivotDataPainterWin32com


logger = logging.getLogger( 'starter')
logger.setLevel(logging.DEBUG)

ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
logger.addHandler(ch)

crm_file_path = "F:\\python\\compareCRMAndSAP\\color1\\CRM.xlsx"
sap_file_path = "F:\\python\\compareCRMAndSAP\\color1\\SAP.xlsx"

crm_result_file_path = "F:\\python\\compareCRMAndSAP\\color1\\crm_result.xlsx"
sap_result_file_path = "F:\\python\\compareCRMAndSAP\\color1\\sap_result.xlsx"

crm_sheet_name = "By Detailed Product"
sap_sheet_name = "Summary by Product"


data_reader = PivotDataReader(file_path=crm_file_path,
                              first_row=9, # actual row - 1
                              max_row = 432, # actual row
                              max_col=32, # actual col
                              skip_cols=[6, 7, 32, 33])

crm_data_container = data_reader.read_sheet_data(crm_sheet_name)

data_reader = PivotDataReader(file_path=sap_file_path,
                              first_row=8,
                              max_row = 198,
                              max_col=22,
                              skip_cols=[6, 7, 22, 23])

sap_data_container = data_reader.read_sheet_data(sap_sheet_name)

data_comparator = PivotDataComparator(crm_data_container, sap_data_container)

data_comparator.compare()

pivot_data_painter = PivotDataPainterWin32com(crm_file_path=crm_file_path,
                                              sap_file_path=sap_file_path,
                                              crm_result_file_path=crm_result_file_path,
                                              sap_result_file_path=sap_result_file_path,
                                              crm_sheet_name=crm_sheet_name,
                                              sap_sheet_name=sap_sheet_name,
                                              crm_data_container=crm_data_container,
                                              sap_data_container=sap_data_container)


pivot_data_painter.save()

print("success")


