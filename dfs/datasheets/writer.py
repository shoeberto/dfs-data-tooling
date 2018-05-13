from openpyxl import Workbook
import dfs.datasheets.datasheet as datasheet


class DatasheetWriter:
    def write(self, sheet, output_directory):
        """
        Write a datasheet to an xlsx file.

        Keyword arguments:
        sheet -- a datasheet object.
        output_directory -- the directory to write the new xlsx file to.
        """

        workbook = Workbook()

        # remove the default worksheet
        workbook.remove(workbook.get_sheet_by_name('Sheet'))

        general_tab = workbook.create_sheet(title=datasheet.TAB_NAME_GENERAL)
        self.format_general_tab(sheet, general_tab)

        workbook.save('{}/{}'.format(output_directory, sheet.input_filename))


    def format_general_tab(self, sheet, tab):
        tab['A1'] = 'Study Area'
        tab['B1'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].study_area

        tab['A2'] = 'Plot Number'
        tab['B2'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].plot_number