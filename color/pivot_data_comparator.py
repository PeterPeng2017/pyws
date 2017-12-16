import logging


from compare.app_constants import AppConstants

logger = logging.getLogger( 'SalesDataReader')
logger.setLevel(logging.DEBUG)

ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
logger.addHandler(ch)

class PivotDataComparator:
    def __init__(self, crm_data_container, sap_data_container):
        self.crm_data_container = crm_data_container
        self.sap_data_container = sap_data_container

    def compare(self):
        crm_sheet_data = self.crm_data_container.sheet_data
        sap_sheet_data = self.sap_data_container.sheet_data

        for crm_row_key, crm_row_data in crm_sheet_data.items():
            if crm_row_key not in sap_sheet_data:
                self.color_only_crm(crm_row_data)
            else:
                sap_row_data = sap_sheet_data[crm_row_key]
                self.compare_row_data(crm_row_data, sap_row_data)

        for sap_row_key, sap_row_data in sap_sheet_data.items():
            if sap_row_key not in crm_sheet_data:
                self.color_only_sap(sap_row_data)

    @staticmethod
    def color_only_crm(crm_row_data):
        for cell_key, cell_unit in crm_row_data.items():
            if len(str(cell_unit.value)) > 0:
                cell_unit.color = AppConstants.YELLOW

    @staticmethod
    def color_only_sap(sap_row_data):
        for cell_key, cell_unit in sap_row_data.items():
            if len(str(cell_unit.value)) > 0:
                cell_unit.color = AppConstants.RED

    @staticmethod
    def compare_row_data(crm_row_data, sap_row_data):
        for crm_cell_key, crm_cell_unit in crm_row_data.items():
            if crm_cell_key in sap_row_data:
                sap_cell_unit = sap_row_data[crm_cell_key]
                sap_cell_value = sap_cell_unit.value
                crm_cell_value = crm_cell_unit.value

                if len(str(crm_cell_value)) > 0 and len(str(sap_cell_value)) > 0:
                    if crm_cell_value == sap_cell_value:
                        crm_cell_unit.color = AppConstants.GREEN
                        sap_cell_unit.color = AppConstants.GREEN
                    else:
                        crm_cell_unit.color = AppConstants.BLUE
                        sap_cell_unit.color = AppConstants.BLUE
                elif len(str(crm_cell_value)) > 0 and len(str(sap_cell_value)) == 0:
                        crm_cell_unit.color = AppConstants.YELLOW
                elif len(str(crm_cell_value)) == 0 and len(str(sap_cell_value)) > 0:
                        sap_cell_unit.color = AppConstants.RED

            else:
                if len(str(crm_cell_unit.value)) > 0:
                    crm_cell_unit.color = AppConstants.YELLOW


