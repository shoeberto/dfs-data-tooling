from dfs.datasheets.parsers.plots.parser_2014 import DatasheetParser2014
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class SuperplotDatasheetParser2014(DatasheetParser2014):
    def format_output_filename(self, input_filename):
        return input_filename.replace('QC', 'QC_Converted')


    def validate_workbook(self, workbook, filepath):
        for t in [datasheet.TAB_NAME_GENERAL, datasheet.TAB_NAME_TREE_TABLE, datasheet.TAB_NAME_COVER_TABLE, datasheet.TAB_NAME_SAPLING, datasheet.TAB_NAME_SEEDLING]:
            if t not in workbook:
                raise Exception("Missing required worksheet '{}' in file '{}'".format(t, filepath))


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.general.SuperplotGeneralTab()

        if worksheet['D3'].value:
            tab.study_area = self.parse_int(worksheet['D3'].value)
            tab.plot_number = self.parse_int(worksheet['D4'].value)

        tab.deer_impact = self.parse_int(worksheet['D5'].value)
        tab.collection_date = worksheet['D6'].value

        tab.fenced_subplot_condition = []

        for rownumber in range(16, 28):
            if None == self.parse_int(worksheet[f'B{rownumber}'].value):
                continue

            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'B{rownumber}'].value)

            subplot.slope = self.parse_int(worksheet[f'C{rownumber}'].value)
            subplot.forested = self.parse_int(worksheet[f'D{rownumber}'].value)

            tab.subplots.append(subplot)

        for rownumber in range(32, 40):
            if None != worksheet[f'C{rownumber}'].value:
                auxillary_post_location = datatabs.general.AuxillaryPostLocation()

                auxillary_post_location.post = self.parse_int(worksheet[f'B{rownumber}'].value.replace('Post ', ''))
                auxillary_post_location.micro_plot_id = self.parse_int(worksheet[f'C{rownumber}'].value)
                auxillary_post_location.azimuth = self.parse_int(worksheet[f'D{rownumber}'].value)
                auxillary_post_location.distance = self.parse_float(worksheet[f'E{rownumber}'].value)

                tab.auxillary_post_locations.append(auxillary_post_location)

        return tab


    def parse_witness_tree_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.witnesstree.WitnessTreeTab()

        for rownumber in range(10, 13):
            tree = datatabs.witnesstree.WitnessTreeTabTree()

            tree.micro_plot_id = 1

            tree.tree_number = self.parse_int(worksheet[f'B{rownumber}'].value[1])
            tree.species_known = worksheet[f'C{rownumber}'].value
            tree.species_guess = worksheet[f'D{rownumber}'].value
            tree.dbh = self.parse_float(worksheet[f'E{rownumber}'].value)

            live_or_dead = worksheet[f'F{rownumber}'].value

            # normalize L/D to uppercase if present, else set the raw value and validate later
            if None != live_or_dead and live_or_dead.lower() in ['l', 'd']:
                tree.live_or_dead = live_or_dead.upper()
            else:
                tree.live_or_dead = live_or_dead

            tree.azimuth = self.parse_int(worksheet[f'G{rownumber}'].value)
            tree.distance = self.parse_int(worksheet[f'H{rownumber}'].value)

            tab.witness_trees.append(tree)
                        
        return tab


    def parse_tree_table_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_TREE_TABLE]
        tab = datatabs.tree.TreeTableTab()

        row_valid = True
        i = 3

        subplot_tree_numbers = {}

        while (row_valid):
            if not worksheet[f'A{i}'].value:
                row_valid = False
                continue

            species = datatabs.tree.TreeTableSpecies()

            species.micro_plot_id = worksheet[f'A{i}'].value

            if species.micro_plot_id not in subplot_tree_numbers:
                subplot_tree_numbers[species.micro_plot_id] = 1
            else:
                subplot_tree_numbers[species.micro_plot_id] += 1

            species.tree_number = subplot_tree_numbers[species.micro_plot_id]
            species.species_known = worksheet[f'C{i}'].value
            species.species_guess = worksheet[f'D{i}'].value
            species.diameter_breast_height = self.parse_float(worksheet[f'E{i}'].value)

            live_or_dead = str(worksheet[f'F{i}'].value).upper()

            if None != live_or_dead and live_or_dead.lower() in ['l', 'd']:
                species.live_or_dead = live_or_dead.upper()
            else:
                species.live_or_dead = live_or_dead

            species.comments = ''

            tab.tree_species.append(species)

            i += 1

        return tab


    def parse_sapling_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_SAPLING]
        tab = datatabs.sapling.SaplingTab()

        row_valid = True
        i = 5

        subplot_sapling_numbers = {}

        while (row_valid):
            if not worksheet[f'A{i}'].value:
                row_valid = False
                continue

            species = datatabs.sapling.SaplingSpecies()

            species.micro_plot_id = self.parse_int(worksheet[f'A{i}'].value)

            if species.micro_plot_id not in subplot_sapling_numbers:
                subplot_sapling_numbers[species.micro_plot_id] = 1
            else:
                subplot_sapling_numbers[species.micro_plot_id] += 1

            species.sapling_number = subplot_sapling_numbers[species.micro_plot_id]
            species.quarter = self.parse_int(worksheet[f'B{i}'].value)
            species.scale = self.parse_int(worksheet[f'C{i}'].value)
            species.species_known = worksheet[f'D{i}'].value
            species.species_guess = worksheet[f'E{i}'].value
            species.diameter_breast_height = self.parse_float(worksheet[f'F{i}'].value)

            tab.sapling_species.append(species)

            i += 1

        return tab


    def parse_seedling_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_SEEDLING]
        tab = datatabs.seedling.SeedlingTable()

        row_valid = True
        i = 5

        while (row_valid):
            if not worksheet[f'A{i}'].value:
                row_valid = False
                continue

            species = datatabs.seedling.SeedlingSpecies()

            species.micro_plot_id = worksheet[f'A{i}'].value

            species.quarter = self.parse_int(worksheet[f'B{i}'].value)
            species.scale = self.parse_int(worksheet[f'C{i}'].value)
            species.species_known = worksheet[f'D{i}'].value
            species.species_guess = worksheet[f'E{i}'].value
            species.sprout = self.parse_int(worksheet[f'F{i}'].value)
            species.zero_six_inches = self.parse_int(worksheet[f'G{i}'].value)
            species.six_twelve_inches = self.parse_int(worksheet[f'H{i}'].value)
            species.one_three_feet_total = self.parse_int(worksheet[f'I{i}'].value)
            species.one_three_feet_browsed = self.parse_int(worksheet[f'J{i}'].value)

            if species.one_three_feet_total and not species.one_three_feet_browsed:
                species.one_three_feet_browsed = 0

            species.three_five_feet_total = self.parse_int(worksheet[f'K{i}'].value)
            species.three_five_feet_browsed = self.parse_int(worksheet[f'L{i}'].value)

            if species.three_five_feet_total and not species.three_five_feet_browsed:
                species.three_five_feet_browsed = 0

            species.greater_five_feet_total = self.parse_int(worksheet[f'M{i}'].value)
            species.greater_five_feet_browsed = self.parse_int(worksheet[f'N{i}'].value)

            if species.greater_five_feet_total and not species.greater_five_feet_browsed:
                species.greater_five_feet_browsed = 0

            tab.seedling_species.append(species)

            i += 1

        return tab

