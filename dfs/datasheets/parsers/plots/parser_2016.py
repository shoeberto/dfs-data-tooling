from dfs.datasheets.parser import DatasheetParser
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class DatasheetParser2016(DatasheetParser):
    def format_output_filename(self, input_filename):
        return input_filename.replace('Data', 'Data_Converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]

        if worksheet['D2'].value == 'Plot Center GPS':
            raise Exception('File does not conform to canonical input format.')

        tab = datatabs.general.PlotGeneralTab()

        tab.study_area = self.parse_int(worksheet['B1'].value)
        tab.plot_number = self.parse_int(worksheet['B2'].value)
        tab.deer_impact = self.parse_int(worksheet['B3'].value)
        tab.collection_date = worksheet['B4'].value

        collected_cover_subplots = self.get_collected_cover_subplots(workbook)

        for rownumber in range(8, 13):
            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

            # Ignore slope
            subplot.latitude = self.parse_float(worksheet[f'K{rownumber}'].value)
            subplot.longitude = self.parse_float(worksheet[f'L{rownumber}'].value)

            subplot.forested = worksheet[f'D{rownumber}'].value

            if None == subplot.forested:
                if subplot.micro_plot_id in collected_cover_subplots or (None != subplot.latitude or None != subplot.longitude):
                    subplot.forested = 'Yes'
                else:
                    subplot.forested = 'No'

            subplot.disturbance = self.parse_int(worksheet[f'E{rownumber}'].value)

            if None == subplot.disturbance:
                subplot.disturbance = 0

            subplot.disturbance_type = self.parse_int(worksheet[f'F{rownumber}'].value)

            if None == subplot.disturbance_type:
                subplot.disturbance_type = 0

            subplot.collected = worksheet[f'G{rownumber}'].value

            if None == subplot.collected:
                if subplot.micro_plot_id in collected_cover_subplots:
                    subplot.collected = 'Yes'
                else:
                    subplot.collected = 'No'

            subplot.fenced = worksheet[f'H{rownumber}'].value

            if None == subplot.fenced:
                if 5 == subplot.micro_plot_id and worksheet['A16'].value:
                    subplot.fenced = 'Yes'
                else:
                    subplot.fenced = 'No'

            subplot.azimuth = self.parse_int(worksheet[f'I{rownumber}'].value)

            if None == subplot.azimuth and subplot.collected:
                if 2 == subplot.micro_plot_id:
                    subplot.azimuth = 0
                elif 3 == subplot.micro_plot_id:
                    subplot.azimuth = 120
                elif 4 == subplot.micro_plot_id:
                    subplot.azimuth = 240
                elif 5 == subplot.micro_plot_id:
                    subplot.azimuth = 60

            subplot.altitude = self.parse_float(worksheet[f'J{rownumber}'].value)

            tab.subplots.append(subplot)


        tab.fenced_subplot_condition = datatabs.general.FencedSubplotConditions()
        tab.fenced_subplot_condition.active_exclosure = worksheet[f'A16'].value
        tab.fenced_subplot_condition.repairs = worksheet[f'B16'].value
        tab.fenced_subplot_condition.level = self.parse_int(worksheet[f'C16'].value)
        tab.fenced_subplot_condition.notes = worksheet[f'D16'].value


        for rownumber in range(40, 45):
            if None != self.parse_int(worksheet[f'A{rownumber}'].value):
                non_forested_azimuth = datatabs.general.NonForestedAzimuths()

                non_forested_azimuth.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

                non_forested_azimuth.azimuth_1 = self.parse_int(worksheet[f'B{rownumber}'].value)
                non_forested_azimuth.azimuth_2 = self.parse_int(worksheet[f'C{rownumber}'].value)
                non_forested_azimuth.azimuth_3 = self.parse_int(worksheet[f'D{rownumber}'].value)


        for rownumber in range(26, 36):
            if None != worksheet[f'C{rownumber}'].value:
                auxillary_post_location = datatabs.general.AuxillaryPostLocation()

                auxillary_post_location.post = self.parse_int(worksheet[f'A{rownumber}'].value.replace('Post ', ''))
                auxillary_post_location.micro_plot_id = self.parse_int(worksheet[f'B{rownumber}'].value)
                auxillary_post_location.stake_type = worksheet[f'C{rownumber}'].value
                auxillary_post_location.azimuth = self.parse_int(worksheet[f'D{rownumber}'].value)
                auxillary_post_location.distance = self.parse_float(worksheet[f'E{rownumber}'].value)

                tab.auxillary_post_locations.append(auxillary_post_location)

        return tab


    def parse_witness_tree_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.witnesstree.WitnessTreeTab()

        for rownumber in range(19, 22):
            tree = datatabs.witnesstree.WitnessTreeTabTree()

            tree.micro_plot_id = 1

            tree.tree_number = self.parse_int(worksheet[f'A{rownumber}'].value[1])
            tree.species_known = worksheet[f'B{rownumber}'].value
            tree.species_guess = worksheet[f'C{rownumber}'].value
            tree.dbh = self.parse_float(worksheet[f'D{rownumber}'].value)

            live_or_dead = self.parse_int(worksheet[f'E{rownumber}'].value)

            if None != live_or_dead:
                tree.live_or_dead = 'L' if 1 == live_or_dead else 'D'
            else:
                tree.live_or_dead = 'L'

            tree.azimuth = self.parse_int(worksheet[f'F{rownumber}'].value)
            tree.distance = self.parse_float(worksheet[f'G{rownumber}'].value)

            if None != tree.species_known:
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

        tab.seedlings = worksheet['B10'].value
        tab.seedlings_notes = self.parse_note_fields(worksheet, 'D10:K11')

        tab.browsing = worksheet['B13'].value
        tab.browsing_notes = self.parse_note_fields(worksheet, 'D13:K14')

        tab.indicators = worksheet['B16'].value
        tab.indicators_notes = self.parse_note_fields(worksheet, 'D16:K17')

        tab.general_notes = self.parse_note_fields(worksheet, 'A21:K41')

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
            else:
                species.live_or_dead = 'L'

            species.comments = worksheet[f'G{i}'].value

            tab.tree_species.append(species)

            i += 1

        return tab


    def get_collected_cover_subplots(self, workbook):
        subplots = []

        worksheet = workbook[datasheet.TAB_NAME_COVER_TABLE]

        row_valid = True
        i = 3

        while (row_valid):
            if not worksheet[f'A{i}'].value:
                row_valid = False
                continue

            subplots.append(self.parse_int(worksheet[f'A{i}'].value))

            i = i + 1

        return list(set(subplots))