from openpyxl import load_workbook
from os.path import basename
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet
from dfs.datasheets.datatabs.tabs import FieldValidationException
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

        # not recorded for 2013
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
                raise FieldValidationException('PlotGeneralTab', 'forested', '0 or 1', forested_value)

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
        worksheet = workbook[datasheet.TAB_NAME_COVER_TABLE]
        tab = datatabs.cover.CoverTableTab()

        row_valid = True
        i = 3

        while (row_valid):
            if not worksheet['A{}'.format(i)].value:
                row_valid = False
                continue

            species = datatabs.cover.CoverSpecies()

            species.micro_plot_id = worksheet['A{}'.format(i)].value
            species.quarter = int(worksheet['B{}'.format(i)].value)
            species.scale = int(worksheet['C{}'.format(i)].value)
            species.species_known = worksheet['D{}'.format(i)].value
            species.species_guess = worksheet['E{}'.format(i)].value
            species.percent_cover = int(worksheet['F{}'.format(i)].value)
            species.average_height = worksheet['G{}'.format(i)].value
            species.count = worksheet['H{}'.format(i)].value
            species.flower = worksheet['I{}'.format(i)].value
            species.number_of_stems = worksheet['J{}'.format(i)].value

            if species.count and not species.flower:
                species.flower = 0

            tab.cover_species.append(species)

            i += 1

        return tab



    def parse_notes_tab(self, workbook):
        # TODO: implement
        return


    def parse_sapling_tab(self, workbook):
        worksheet = workbook[datasheet.TAB_NAME_SAPLING]
        tab = datatabs.sapling.SaplingTab()

        row_valid = True
        i = 3

        subplot_sapling_numbers = {}

        while (row_valid):
            if not worksheet['A{}'.format(i)].value:
                row_valid = False
                continue

            species = datatabs.sapling.SaplingSpecies()

            species.micro_plot_id = worksheet['A{}'.format(i)].value

            if species.micro_plot_id not in subplot_sapling_numbers:
                subplot_sapling_numbers[species.micro_plot_id] = 1
            else:
                subplot_sapling_numbers[species.micro_plot_id] += 1

            species.sapling_number = subplot_sapling_numbers[species.micro_plot_id]
            species.quarter = int(worksheet['B{}'.format(i)].value)
            species.scale = int(worksheet['C{}'.format(i)].value)
            species.species_known = worksheet['D{}'.format(i)].value
            species.species_guess = worksheet['E{}'.format(i)].value
            species.diameter_breast_height = float(worksheet['F{}'.format(i)].value)

            tab.sapling_species.append(species)

            i += 1

        return tab


    def parse_seedling_tab(self, workbook):
        worksheet = workbook[datasheet.TAB_NAME_SEEDLING]
        tab = datatabs.seedling.SeedlingTable()

        row_valid = True
        i = 3

        while (row_valid):
            if not worksheet['A{}'.format(i)].value:
                row_valid = False
                continue

            species = datatabs.seedling.SeedlingSpecies()

            species.micro_plot_id = worksheet['A{}'.format(i)].value

            species.quarter = int(worksheet['B{}'.format(i)].value)
            species.scale = int(worksheet['C{}'.format(i)].value)
            species.species_known = worksheet['D{}'.format(i)].value
            species.species_guess = worksheet['E{}'.format(i)].value
            species.sprout = 0
            species.zero_six_inches = worksheet['F{}'.format(i)].value
            species.six_twelve_inches = worksheet['G{}'.format(i)].value
            species.one_three_feet_total = worksheet['H{}'.format(i)].value
            species.one_three_feet_browsed = worksheet['I{}'.format(i)].value

            if species.one_three_feet_total and not species.one_three_feet_browsed:
                species.one_three_feet_browsed = 0

            species.three_five_feet_total = worksheet['J{}'.format(i)].value
            species.three_five_feet_browsed = worksheet['K{}'.format(i)].value

            if species.three_five_feet_total and not species.three_five_feet_browsed:
                species.three_five_feet_browsed = 0

            species.greater_five_feet_total = worksheet['L{}'.format(i)].value
            species.greater_five_feet_browsed = worksheet['M{}'.format(i)].value

            if species.greater_five_feet_total and not species.greater_five_feet_browsed:
                species.greater_five_feet_browsed = 0

            tab.seedling_species.append(species)

            i += 1

        return tab


    def parse_tree_table_tab(self, workbook):
        worksheet = workbook[datasheet.TAB_NAME_TREE_TABLE]
        tab = datatabs.tree.TreeTableTab()

        row_valid = True
        i = 3

        subplot_tree_numbers = {}

        while (row_valid):
            if not worksheet['A{}'.format(i)].value:
                row_valid = False
                continue

            species = datatabs.tree.TreeTableSpecies()

            species.micro_plot_id = worksheet['A{}'.format(i)].value

            if species.micro_plot_id not in subplot_tree_numbers:
                subplot_tree_numbers[species.micro_plot_id] = 1
            else:
                subplot_tree_numbers[species.micro_plot_id] += 1

            species.tree_number = subplot_tree_numbers[species.micro_plot_id]
            species.species_known = worksheet['C{}'.format(i)].value
            species.species_guess = worksheet['D{}'.format(i)].value
            species.diameter_breast_height = float(worksheet['E{}'.format(i)].value)
            species.live_or_dead = worksheet['F{}'.format(i)].value
            species.comments = ''

            tab.tree_species.append(species)

            i += 1

        return tab

