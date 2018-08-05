from openpyxl import load_workbook
from os.path import basename
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet
from dfs.datasheets.datatabs.tabs import FieldValidationError
from abc import ABC, abstractmethod


class DatasheetParser(ABC):
    def parse_datasheet(self, filepath):
        """
        Parse a datasheet file.

        Keyword arguments:
        filepath -- full path to the datasheet.
        """

        workbook = load_workbook(filepath)

        self.validate_workbook(workbook, filepath)

        sheet = datasheet.Datasheet()
        sheet.input_filename = basename(filepath)

        sheet.tabs[datasheet.TAB_NAME_GENERAL] = self.parse_plot_general_tab(workbook, sheet)
        sheet.tabs[datasheet.TAB_NAME_WITNESS_TREES] = self.parse_witness_tree_tab(workbook, sheet)
        sheet.tabs[datasheet.TAB_NAME_NOTES] = self.parse_notes_tab(workbook, sheet)
        sheet.tabs[datasheet.TAB_NAME_TREE_TABLE] = self.parse_tree_table_tab(workbook, sheet)
        sheet.tabs[datasheet.TAB_NAME_COVER_TABLE] = self.parse_cover_table_tab(workbook, sheet)
        sheet.tabs[datasheet.TAB_NAME_SAPLING] = self.parse_sapling_tab(workbook, sheet)
        sheet.tabs[datasheet.TAB_NAME_SEEDLING] = self.parse_seedling_tab(workbook, sheet)

        return sheet

    
    def validate_workbook(self, workbook, filepath):
        # by default, validate all sheets that are required across all years
        # this excludes Witness Trees, as they were not present in older datasheets
        for t in [datasheet.TAB_NAME_GENERAL, datasheet.TAB_NAME_NOTES, datasheet.TAB_NAME_TREE_TABLE, datasheet.TAB_NAME_COVER_TABLE, datasheet.TAB_NAME_SAPLING, datasheet.TAB_NAME_SEEDLING]:
            if t not in workbook:
                raise Exception("Missing required worksheet '{}' in file '{}'".format(t, filepath))


    @abstractmethod
    def parse_plot_general_tab(self, workbook, sheet):
        """
        Parse out all data composing the "General" tab.
        
        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_witness_tree_tab(self, workbook, sheet):
        """
        Parse out all data composing the "Witness_Trees" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_notes_tab(self, workbook, sheet):
        """
        Parse out all data composing the "Notes" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_tree_table_tab(self, workbook, sheet):
        """
        Parse out all data composing the "Tree_Table" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_cover_table_tab(self, workbook, sheet):
        """
        Parse out all data composing the "Cover_Table" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_sapling_tab(self, workbook, sheet):
        """
        Parse out all data composing the "Sapling_(1-5)" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_seedling_tab(self, workbook, sheet):
        """
        Parse out all data composing the "Seedling_(0-1)" tab.

        Keyword arguments:
        workbook -- the source workbook
        """

    
    def parse_float(self, record):
        if None == record:
            return None

        if '' == record:
            return None

        try:
            return float(record)
        except Exception:
            return None

    
    def parse_int(self, record):
        if None == record:
            return None

        if '' == record:
            return None

        if isinstance(record, float):
            return int(round(record))

        try:
            return int(record)
        except Exception:
            return None