from dfs.datasheets.parsers.plots.parser_2017 import DatasheetParser2017
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class TreatmentDatasheetParser2014(DatasheetParser2017):
    def format_output_filename(self, input_filename):
        return input_filename.replace('Merged', 'Merged_Converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]

        tab = datatabs.general.GeneralTab()

        tab.study_area = self.parse_int(worksheet['B1'].value)
        tab.plot_number = self.parse_int(worksheet['B2'].value)
        tab.deer_impact = self.parse_int(workbook[datasheet.TAB_NAME_NOTES]['C3'].value)
        tab.collection_date = worksheet['B4'].value

        for rownumber in range(9, 24):
            subplot = datatabs.general.TreatmentPlotGeneralPlotSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'C{rownumber}'].value)

            # Ignore slope
            subplot.latitude = self.parse_float(worksheet[f'A{rownumber}'].value)
            subplot.longitude = self.parse_float(worksheet[f'B{rownumber}'].value)

            subplot.collected = worksheet[f'D{rownumber}'].value
            subplot.re_monumented = worksheet[f'E{rownumber}'].value
            subplot.forested = worksheet[f'F{rownumber}'].value
            subplot.disturbance = self.parse_int(worksheet[f'G{rownumber}'].value)
            subplot.disturbance_type = self.parse_int(worksheet[f'H{rownumber}'].value)
            subplot.altitude = self.parse_float(worksheet[f'I{rownumber}'].value)
            subplot.notes = worksheet[f'J{rownumber}'].value

            tab.subplots.append(subplot)


        for rownumber in range(29, 34):
            if None != worksheet[f'A{rownumber}'].value:
                auxillary_post_location = datatabs.general.AuxillaryPostLocation()

                auxillary_post_location.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)
                auxillary_post_location.post = self.parse_int(worksheet[f'B{rownumber}'].value)
                auxillary_post_location.stake_type = worksheet[f'C{rownumber}'].value.capitalize()
                auxillary_post_location.azimuth = self.parse_int(worksheet[f'D{rownumber}'].value)
                auxillary_post_location.distance = self.parse_float(worksheet[f'E{rownumber}'].value)

                tab.auxillary_post_locations.append(auxillary_post_location)


        for rownumber in range(37, 42):
            if None != self.parse_int(worksheet[f'A{rownumber}'].value):
                non_forested_azimuth = datatabs.general.NonForestedAzimuths()

                non_forested_azimuth.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

                non_forested_azimuth.azimuth_1 = self.parse_int(worksheet[f'B{rownumber}'].value)
                non_forested_azimuth.azimuth_2 = self.parse_int(worksheet[f'C{rownumber}'].value)
                non_forested_azimuth.azimuth_3 = self.parse_int(worksheet[f'D{rownumber}'].value)

        return tab


    def parse_witness_tree_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_WITNESS_TREES]
        tab = datatabs.witnesstree.WitnessTreeTab()

        for rownumber in range(3, 33):
            tree = datatabs.witnesstree.WitnessTreeTabTree()

            tree.micro_plot_id = self.parse_int(worksheet[f'B{rownumber}'].value)

            if None == tree.micro_plot_id:
                continue

            tree.tree_number = self.parse_int(worksheet[f'A{rownumber}'].value)

            tree.species_known = worksheet[f'C{rownumber}'].value
            tree.species_guess = worksheet[f'D{rownumber}'].value
            tree.dbh = self.parse_float(worksheet[f'E{rownumber}'].value)

            tree.live_or_dead = worksheet[f'F{rownumber}'].value

            tree.azimuth = self.parse_int(worksheet[f'G{rownumber}'].value)
            tree.distance = self.parse_float(worksheet[f'H{rownumber}'].value)

            tab.witness_trees.append(tree)

        return tab


    def parse_sapling_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_SAPLING]
        tab = datatabs.sapling.SaplingTab()

        row_valid = True
        i = 3

        while (row_valid):
            if not worksheet[f'A{i}'].value:
                row_valid = False
                continue

            species = datatabs.sapling.SaplingSpecies()

            species.micro_plot_id = self.parse_int(worksheet[f'A{i}'].value)

            species.sapling_number = self.parse_int(worksheet[f'B{i}'].value)
            species.quarter = self.parse_int(worksheet[f'C{i}'].value)
            species.scale = self.parse_int(worksheet[f'D{i}'].value)
            species.species_known = worksheet[f'E{i}'].value
            species.species_guess = worksheet[f'F{i}'].value
            species.diameter_breast_height = self.parse_float(worksheet[f'G{i}'].value)

            tab.sapling_species.append(species)

            i += 1

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

            species.live_or_dead = worksheet[f'F{i}'].value

            species.comments = worksheet[f'G{i}'].value

            tab.tree_species.append(species)

            i += 1

        return tab