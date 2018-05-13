from openpyxl import load_workbook
from os.path import basename
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet
import abc


class DatasheetParser:
    __metaclass__ = abc.ABCMeta

    def parse_datasheet(self, filepath):
        """
        Parse a datasheet file.

        Keyword arguments:
        filepath -- full path to the datasheet.
        """

        workbook = load_workbook(filepath)

        self.validate_workbook(workbook)

        sheet = datasheet.Datasheet()
        sheet.input_filename = basename(filepath)

        sheet.tabs[datasheet.TAB_NAME_GENERAL] = self.parse_general_tab(workbook[datasheet.TAB_NAME_GENERAL])

        return sheet

    
    def validate_workbook(self, workbook):
        # by default, validate all sheets that are required across all years
        # this excludes Witness Trees, as they were not present in older datasheets
        for t in [datasheet.TAB_NAME_GENERAL, datasheet.TAB_NAME_NOTES, datasheet.TAB_NAME_TREE_TABLE, datasheet.TAB_NAME_COVER_TABLE, datasheet.TAB_NAME_SAPLING, datasheet.TAB_NAME_SEEDLING]:
            if t not in workbook:
                raise Exception("Missing required worksheet '{}' in file '{}'".format(t, filePath))


    @abc.abstractmethod
    def parse_general_tab(self, worksheet):
        """
        Parse out all data from the "General" tab.
        
        Keyword arguments:
        worksheet -- the "General" worksheet object
        """
        return


class DatasheetParser2013(DatasheetParser):
    def parse_general_tab(self, worksheet):
        tab = datatabs.general.GeneralTab()

        tab.study_area = worksheet['D3'].value
        tab.plot_number = worksheet['D4'].value
        tab.deer_impact = worksheet['D5'].value
        tab.date = worksheet['D6'].value

        for rownumber in range(16, 21):
            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = worksheet['B{}'.format(rownumber)].value
            # Ignore slope
            subplot.forested = worksheet['D{}'.format(rownumber)].value

            tab.subplots.append(subplot)

        return tab