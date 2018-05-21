from openpyxl import load_workbook
from os.path import basename
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet
from abc import ABC, abstractmethod


class DatasheetParser(ABC):
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

        sheet.tabs[datasheet.TAB_NAME_GENERAL] = self.parse_plot_general_tab(workbook)
        sheet.tabs[datasheet.TAB_NAME_WITNESS_TREES] = self.parse_witness_tree_tab(workbook)
        sheet.tabs[datasheet.TAB_NAME_NOTES] = self.parse_notes_tab(workbook)
        sheet.tabs[datasheet.TAB_NAME_TREE_TABLE] = self.parse_tree_table_tab(workbook)
        sheet.tabs[datasheet.TAB_NAME_COVER_TABLE] = self.parse_cover_table_tab(workbook)
        sheet.tabs[datasheet.TAB_NAME_SAPLING] = self.parse_sapling_tab(workbook)
        sheet.tabs[datasheet.TAB_NAME_SEEDLING] = self.parse_seedling_tab(workbook)

        for tab in sheet.tabs.values():
            # TODO: remove this; only useful while some tabs are unimplemented
            if None == tab:
                continue

            tab.validate()

        return sheet

    
    def validate_workbook(self, workbook):
        # by default, validate all sheets that are required across all years
        # this excludes Witness Trees, as they were not present in older datasheets
        for t in [datasheet.TAB_NAME_GENERAL, datasheet.TAB_NAME_NOTES, datasheet.TAB_NAME_TREE_TABLE, datasheet.TAB_NAME_COVER_TABLE, datasheet.TAB_NAME_SAPLING, datasheet.TAB_NAME_SEEDLING]:
            if t not in workbook:
                raise Exception("Missing required worksheet '{}' in file '{}'".format(t, filePath))


    @abstractmethod
    def parse_plot_general_tab(self, workbook):
        """
        Parse out all data composing the "General" tab.
        
        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_witness_tree_tab(self, workbook):
        """
        Parse out all data composing the "Witness_Trees" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_notes_tab(self, workbook):
        """
        Parse out all data composing the "Notes" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_tree_table_tab(self, workbook):
        """
        Parse out all data composing the "Tree_Table" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_cover_table_tab(self, workbook):
        """
        Parse out all data composing the "Cover_Table" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_sapling_tab(self, workbook):
        """
        Parse out all data composing the "Sapling_(1-5)" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


    @abstractmethod
    def parse_seedling_tab(self, workbook):
        """
        Parse out all data composing the "Seedling_(0-1)" tab.

        Keyword arguments:
        workbook -- the source workbook
        """


class DatasheetParser2013(DatasheetParser):
    def parse_plot_general_tab(self, workbook):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.general.PlotGeneralTab()

        tab.study_area = worksheet['D3'].value
        tab.plot_number = worksheet['D4'].value
        tab.deer_impact = worksheet['D5'].value
        tab.collection_date = worksheet['D6'].value

        # TODO not recorded in 2013?
        # leaving blank for now
        tab.fenced_subplot_condition = datatabs.general.FencedSubplotConditions()

        for rownumber in range(16, 21):
            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = worksheet['B{}'.format(rownumber)].value
            # Ignore slope
            forested_value = worksheet['D{}'.format(rownumber)].value

            if 1 == forested_value:
                subplot.forested = 'Yes'
            elif 0 == forested_value:
                subplot.forested = 'No'
            else:
                raise ValueError("Unexpected value for forested: '{}' in file {}".format(forested_value, workbook.input_filename))

            tab.subplots.append(subplot)

        for rownumber in range(25, 33):
            if worksheet['C{}'.format(rownumber)].value:
                auxillary_post_location = datatabs.general.AuxillaryPostLocation()
                auxillary_post_location.subplot = worksheet['C{}'.format(rownumber)].value
                auxillary_post_location.azimuth = worksheet['D{}'.format(rownumber)].value
                auxillary_post_location.distance = worksheet['E{}'.format(rownumber)].value

        return tab


    def parse_witness_tree_tab(self, workbook):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.witnesstree.WitnessTreeTab()

        for rownumber in range(10, 13):
            tree = datatabs.witnesstree.WitnessTreeTabTree()

            # in 2013, all witness trees were at microplot 1
            tree.micro_plot_id = 1

            tree.tree_number = worksheet['B{}'.format(rownumber)].value[1]
            tree.species_known = worksheet['C{}'.format(rownumber)].value
            tree.species_guess = worksheet['D{}'.format(rownumber)].value
            tree.dbh = worksheet['E{}'.format(rownumber)].value
            tree.live_or_dead = worksheet['F{}'.format(rownumber)].value
            tree.azimuth = worksheet['G{}'.format(rownumber)].value
            tree.distance = worksheet['H{}'.format(rownumber)].value

            tab.witness_trees.append(tree)

        return tab


    def parse_cover_table_tab(self, workbook):
        # TODO: implement
        return


    def parse_notes_tab(self, workbook):
        # TODO: implement
        return


    def parse_sapling_tab(self, workbook):
        # TODO: implement
        return


    def parse_seedling_tab(self, workbook):
        # TODO: implement
        return


    def parse_tree_table_tab(self, workbook):
        # TODO: implement
        return