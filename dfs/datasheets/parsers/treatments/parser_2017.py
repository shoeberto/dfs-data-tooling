from dfs.datasheets.parsers.treatments.parser_2014 import TreatmentDatasheetParser2014
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class TreatmentDatasheetParser2017(TreatmentDatasheetParser2014):
    def format_output_filename(self, input_filename):
        return input_filename.replace('2017', '2017_Converted')


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

            live_or_dead = self.parse_int(worksheet[f'F{rownumber}'].value)

            if None != live_or_dead:
                tree.live_or_dead = 'L' if 1 == live_or_dead else 'D'
            else:
                tree.live_or_dead = 'L'

            tree.azimuth = self.parse_int(worksheet[f'G{rownumber}'].value)
            tree.distance = self.parse_float(worksheet[f'H{rownumber}'].value)

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

            live_or_dead = self.parse_int(worksheet[f'F{i}'].value)

            if None != live_or_dead:
                species.live_or_dead = 'L' if 1 == live_or_dead else 'D'

            species.comments = worksheet[f'G{i}'].value

            tab.tree_species.append(species)

            i += 1

        return tab


    def parse_sapling_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_SAPLING]
        tab = datatabs.sapling.SaplingTab()

        row_valid = True
        i = 3

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

