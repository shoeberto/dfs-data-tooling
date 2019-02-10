from dfs.datasheets.parser import DatasheetParser
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class DatasheetParser2013(DatasheetParser):
    def format_output_filename(self, input_filename):
        return input_filename.replace('Data', 'Data_Converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.general.PlotGeneralTab()

        tab.study_area = self.parse_int(worksheet['D3'].value)
        tab.plot_number = self.parse_int(worksheet['D4'].value)
        tab.deer_impact = self.parse_int(worksheet['D5'].value)
        tab.collection_date = worksheet['D6'].value

        # not recorded for 2013
        tab.fenced_subplot_condition = datatabs.general.FencedSubplotConditions()

        for rownumber in range(16, 21):
            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'B{rownumber}'].value)
            # Ignore slope
            forested_value = worksheet[f'D{rownumber}'].value

            if 1 == forested_value:
                subplot.forested = 'Yes'
            elif 0 == forested_value:
                subplot.forested = 'No'
            else:
                pass

            tab.subplots.append(subplot)

        for rownumber in range(25, 33):
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

            # in 2013, all witness trees were at microplot 1
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


    def parse_cover_table_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_COVER_TABLE]
        tab = datatabs.cover.CoverTableTab()

        row_valid = True
        i = 3

        while (row_valid):
            if not worksheet[f'A{i}'].value:
                row_valid = False
                continue

            species = datatabs.cover.CoverSpecies()

            species.micro_plot_id = self.parse_int(worksheet[f'A{i}'].value)
            species.quarter = self.parse_int(worksheet[f'B{i}'].value)
            species.scale = self.parse_int(worksheet[f'C{i}'].value)
            species.species_known = worksheet[f'D{i}'].value
            species.species_guess = worksheet[f'E{i}'].value

            species.flower = self.parse_int(worksheet[f'I{i}'].value)
            species.number_of_stems = self.parse_int(worksheet[f'J{i}'].value)

            if species.species_known in datatabs.cover.CoverSpecies.DEER_INDICATOR_SPECIES:
                if None == species.flower:
                    species.flower = 0

            species.percent_cover = self.parse_int(worksheet[f'F{i}'].value)
            species.average_height = self.parse_int(worksheet[f'G{i}'].value)
            species.count = self.parse_int(worksheet[f'H{i}'].value)

            if species.count and (not species.flower) and species.override_species(species.species_known) in datatabs.cover.CoverSpecies.DEER_INDICATOR_SPECIES:
                species.flower = 0

            tab.cover_species.append(species)

            i += 1

        return tab


    def parse_notes_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_NOTES]
        tab = datatabs.notes.NotesTab()

        tab.seedlings = worksheet['D11'].value
        tab.seedlings_notes = worksheet['F11'].value

        tab.browsing = worksheet['D14'].value
        tab.browsing_notes = worksheet['F14'].value

        tab.indicators = worksheet['D17'].value
        tab.indicators_notes = worksheet['F17'].value

        notes_rows = []
        for row in worksheet['C22:M42']:
            for cell in row:
                if None != cell.value:
                    notes_rows.append(cell.value)

        tab.general_notes = ' '.join(notes_rows).strip()

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


    def parse_seedling_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_SEEDLING]
        tab = datatabs.seedling.SeedlingTable()

        row_valid = True
        i = 3

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
            species.sprout = 0
            species.zero_six_inches = self.parse_int(worksheet[f'F{i}'].value)
            species.six_twelve_inches = self.parse_int(worksheet[f'G{i}'].value)
            species.one_three_feet_total = self.parse_int(worksheet[f'H{i}'].value)
            species.one_three_feet_browsed = self.parse_int(worksheet[f'I{i}'].value)

            if species.one_three_feet_total and not species.one_three_feet_browsed:
                species.one_three_feet_browsed = 0

            species.three_five_feet_total = self.parse_int(worksheet[f'J{i}'].value)
            species.three_five_feet_browsed = self.parse_int(worksheet[f'K{i}'].value)

            if species.three_five_feet_total and not species.three_five_feet_browsed:
                species.three_five_feet_browsed = 0

            species.greater_five_feet_total = self.parse_int(worksheet[f'L{i}'].value)
            species.greater_five_feet_browsed = self.parse_int(worksheet[f'M{i}'].value)

            if species.greater_five_feet_total and not species.greater_five_feet_browsed:
                species.greater_five_feet_browsed = 0

            tab.seedling_species.append(species)

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

            live_or_dead = str(worksheet[f'F{i}'].value).upper()

            if None != live_or_dead and live_or_dead.lower() in ['l', 'd']:
                species.live_or_dead = live_or_dead.upper()
            else:
                species.live_or_dead = live_or_dead



            species.comments = ''

            tab.tree_species.append(species)

            i += 1

        return tab

