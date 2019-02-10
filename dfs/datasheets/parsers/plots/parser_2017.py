from dfs.datasheets.parser import DatasheetParser
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class DatasheetParser2017(DatasheetParser):
    def format_output_filename(self, input_filename):
        return input_filename.replace('Data', 'Data_Converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]

        tab = datatabs.general.PlotGeneralTab()

        tab.study_area = self.parse_int(worksheet['B1'].value)
        tab.plot_number = self.parse_int(worksheet['B2'].value)
        tab.deer_impact = self.parse_int(worksheet['B3'].value)
        tab.collection_date = worksheet['B4'].value

        for rownumber in range(9, 14):
            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'C{rownumber}'].value)

            # Ignore slope
            subplot.latitude = self.parse_float(worksheet[f'A{rownumber}'].value)
            subplot.longitude = self.parse_float(worksheet[f'B{rownumber}'].value)

            subplot.collected = worksheet[f'D{rownumber}'].value
            subplot.fenced = worksheet[f'E{rownumber}'].value
            subplot.azimuth = self.parse_int(worksheet[f'F{rownumber}'].value)
            subplot.distance = self.parse_int(worksheet[f'G{rownumber}'].value)
            subplot.altitude = self.parse_float(worksheet[f'H{rownumber}'].value)

            # offset rownumber to get the next block of data for this subplot
            rownumber_offset = rownumber + 8
            subplot.re_monumented = worksheet[f'B{rownumber_offset}'].value
            subplot.forested = worksheet[f'C{rownumber_offset}'].value
            subplot.disturbance = self.parse_int(worksheet[f'D{rownumber_offset}'].value)
            subplot.disturbance_type = self.parse_int(worksheet[f'E{rownumber_offset}'].value)
            subplot.lime = 'No' if None == worksheet[f'F{rownumber_offset}'].value else worksheet[f'F{rownumber_offset}'].value.strip()
            subplot.herbicide = 'No' if None == worksheet[f'G{rownumber_offset}'].value else worksheet[f'G{rownumber_offset}'].value.strip()
            subplot.notes = worksheet[f'H{rownumber_offset}'].value

            tab.subplots.append(subplot)


        tab.fenced_subplot_condition = datatabs.general.FencedSubplotConditions()
        tab.fenced_subplot_condition.active_exclosure = worksheet[f'A25'].value
        tab.fenced_subplot_condition.repairs = worksheet[f'B25'].value
        tab.fenced_subplot_condition.level = self.parse_int(worksheet[f'C25'].value)
        tab.fenced_subplot_condition.notes = worksheet[f'D25'].value


        for rownumber in range(44, 49):
            if None != worksheet[f'A{rownumber}'].value:
                auxillary_post_location = datatabs.general.AuxillaryPostLocation()

                auxillary_post_location.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)
                auxillary_post_location.post = self.parse_int(worksheet[f'B{rownumber}'].value)
                auxillary_post_location.stake_type = worksheet[f'C{rownumber}'].value.capitalize()
                auxillary_post_location.azimuth = self.parse_int(worksheet[f'D{rownumber}'].value)
                auxillary_post_location.distance = self.parse_float(worksheet[f'E{rownumber}'].value)

                tab.auxillary_post_locations.append(auxillary_post_location)


        for rownumber in range(52, 57):
            if None != self.parse_int(worksheet[f'A{rownumber}'].value):
                non_forested_azimuth = datatabs.general.NonForestedAzimuths()

                non_forested_azimuth.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

                non_forested_azimuth.azimuth_1 = self.parse_int(worksheet[f'B{rownumber}'].value)
                non_forested_azimuth.azimuth_2 = self.parse_int(worksheet[f'C{rownumber}'].value)
                non_forested_azimuth.azimuth_3 = self.parse_int(worksheet[f'D{rownumber}'].value)

        return tab


    def parse_witness_tree_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.witnesstree.WitnessTreeTab()

        for rownumber in range(29, 40):
            tree = datatabs.witnesstree.WitnessTreeTabTree()

            tree.tree_number = self.parse_int(worksheet[f'A{rownumber}'].value)
            tree.micro_plot_id = self.parse_int(worksheet[f'B{rownumber}'].value)

            tree.species_known = worksheet[f'C{rownumber}'].value
            tree.species_guess = worksheet[f'D{rownumber}'].value
            tree.dbh = self.parse_float(worksheet[f'E{rownumber}'].value)

            live_or_dead = self.parse_int(worksheet[f'F{rownumber}'].value)

            if None != live_or_dead:
                tree.live_or_dead = 'L' if 1 == live_or_dead else 'D'

            tree.azimuth = self.parse_int(worksheet[f'G{rownumber}'].value)
            tree.distance = self.parse_float(worksheet[f'H{rownumber}'].value)

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

            species.percent_cover = self.parse_int(worksheet[f'F{i}'].value)
            species.average_height = self.parse_int(worksheet[f'G{i}'].value)
            species.count = self.parse_int(worksheet[f'H{i}'].value)

            species.flower = self.parse_int(worksheet[f'I{i}'].value)
            species.number_of_stems = self.parse_int(worksheet[f'J{i}'].value)

            if species.species_known in datatabs.cover.CoverSpecies.DEER_INDICATOR_SPECIES:
                if None == species.flower:
                    species.flower = 0

            if species.count and (not species.flower) and species.override_species(species.species_known) in datatabs.cover.CoverSpecies.DEER_INDICATOR_SPECIES:
                species.flower = 0

            tab.cover_species.append(species)

            i += 1

        return tab


    def parse_notes_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_NOTES]
        tab = datatabs.notes.NotesTab()

        tab.seedlings = worksheet['B8'].value
        tab.seedlings_notes = worksheet['D8'].value

        tab.browsing = worksheet['B11'].value
        tab.browsing_notes = worksheet['D11'].value

        tab.indicators = worksheet['B14'].value
        tab.indicators_notes = worksheet['D14'].value

        tab.general_notes = self.parse_note_fields(worksheet, 'A19:K29')

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
            species.sprout = self.parse_int(worksheet[f'F{i}'].value)

            if None == species.sprout:
                species.sprout = 0

            species.zero_six_inches = worksheet[f'G{i}'].value
            species.six_twelve_inches = worksheet[f'H{i}'].value
            species.one_three_feet_total = worksheet[f'I{i}'].value
            species.one_three_feet_browsed = worksheet[f'J{i}'].value

            if species.one_three_feet_total and not species.one_three_feet_browsed:
                species.one_three_feet_browsed = 0

            species.three_five_feet_total = worksheet[f'K{i}'].value
            species.three_five_feet_browsed = worksheet[f'L{i}'].value

            if species.three_five_feet_total and not species.three_five_feet_browsed:
                species.three_five_feet_browsed = 0

            species.greater_five_feet_total = worksheet[f'M{i}'].value
            species.greater_five_feet_browsed = worksheet[f'N{i}'].value

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

            live_or_dead = self.parse_int(worksheet[f'F{i}'].value)

            if None != live_or_dead:
                species.live_or_dead = 'L' if 1 == live_or_dead else 'D'

            species.comments = worksheet[f'G{i}'].value

            tab.tree_species.append(species)

            i += 1

        return tab

