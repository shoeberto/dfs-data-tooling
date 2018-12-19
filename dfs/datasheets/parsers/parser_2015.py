from dfs.datasheets.parser import DatasheetParser
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class DatasheetParser2015(DatasheetParser):
    SUBPLOT_FORESTED_OVERRIDES = {
        116: { 2: False },
        239: { 3: False, 4: False },
        350: { 4: False },
        311: { 4: True },
        313: { 4: True },
        317: { 2: True, 4: True, 5: True },
        325: { 1: True, 3: True, 5: True },
        338: { 3: True },
        426: { 1: True, 3: True, 5: True },
        428: { 3: False },
        431: { 2: True, 4: True, 5: True },
        436: { 3: True },
        438: { 2: False }
    }


    def format_output_filename(self, input_filename):
        return input_filename.replace('data', 'data_converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.general.PlotGeneralTab()

        tab.study_area = self.parse_int(worksheet['B1'].value)
        tab.plot_number = self.parse_int(worksheet['B2'].value)
        tab.deer_impact = self.parse_int(worksheet['B3'].value)
        tab.collection_date = worksheet['B4'].value

        # ignore lat/lon at top of file

        # not recorded for 2015
        tab.fenced_subplot_condition = datatabs.general.FencedSubplotConditions()

        collected_cover_subplots = self.get_collected_cover_subplots(workbook)

        for rownumber in range(27, 42):
            latitude = self.parse_float(worksheet[f'B{rownumber}'].value)
            longitude = self.parse_float(worksheet[f'C{rownumber}'].value)
            micro_plot_id = self.parse_int(worksheet['A{}'.format(rownumber)].value)

            if micro_plot_id > 5 and (None == latitude or None == longitude):
                continue

            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = micro_plot_id

            # Ignore slope
            subplot.converted_latitude = latitude
            subplot.converted_longitude = longitude

            subplot.disturbance = self.parse_int(worksheet[f'D{rownumber}'].value)
            subplot.disturbance_type = self.parse_int(worksheet[f'E{rownumber}'].value)

            subplot.forested = 'Yes' if subplot.micro_plot_id in collected_cover_subplots else None

            if None == subplot.forested:
                plot_id = int(str(tab.study_area) + str(tab.plot_number))
                if plot_id in DatasheetParser2015.SUBPLOT_FORESTED_OVERRIDES and subplot.micro_plot_id in DatasheetParser2015.SUBPLOT_FORESTED_OVERRIDES[plot_id]:
                    subplot.forested = 'Yes' if DatasheetParser2015.SUBPLOT_FORESTED_OVERRIDES[plot_id][subplot.micro_plot_id] else 'No'


            tab.subplots.append(subplot)

        for rownumber in range(46, 51):
            if None != self.parse_int(worksheet[f'A{rownumber}'].value):
                non_forested_azimuth = datatabs.general.NonForestedAzimuths()

                non_forested_azimuth.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

                non_forested_azimuth.azimuth_1 = self.parse_int(worksheet[f'B{rownumber}'].value)
                non_forested_azimuth.azimuth_2 = self.parse_int(worksheet[f'C{rownumber}'].value)
                non_forested_azimuth.azimuth_3 = self.parse_int(worksheet[f'D{rownumber}'].value)


        for rownumber in range(15, 25):
            if None != worksheet[f'C{rownumber}'].value:
                auxillary_post_location = datatabs.general.AuxillaryPostLocation()

                auxillary_post_location.post = self.parse_int(worksheet[f'A{rownumber}'].value.replace('Post ', ''))
                auxillary_post_location.micro_plot_id = self.parse_int(worksheet[f'B{rownumber}'].value)
                auxillary_post_location.azimuth = self.parse_int(worksheet[f'D{rownumber}'].value)
                auxillary_post_location.distance = self.parse_float(worksheet[f'E{rownumber}'].value)

                tab.auxillary_post_locations.append(auxillary_post_location)

        return tab


    def parse_witness_tree_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.witnesstree.WitnessTreeTab()

        for rownumber in range(8, 11):
            tree = datatabs.witnesstree.WitnessTreeTabTree()

            tree.micro_plot_id = 1

            tree.tree_number = self.parse_int(worksheet[f'A{rownumber}'].value[1])
            tree.species_known = worksheet[f'B{rownumber}'].value
            tree.species_guess = worksheet[f'C{rownumber}'].value
            tree.dbh = self.parse_float(worksheet[f'D{rownumber}'].value)

            live_or_dead = self.parse_int(worksheet[f'E{rownumber}'].value)

            if None != live_or_dead:
                tree.live_or_dead = 'L' if 1 == live_or_dead else 'D'

            tree.azimuth = self.parse_int(worksheet[f'F{rownumber}'].value)
            tree.distance = self.parse_float(worksheet[f'G{rownumber}'].value)

            tab.witness_trees.append(tree)
                        
        return tab


    def parse_cover_table_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_COVER_TABLE]
        tab = datatabs.cover.CoverTableTab()

        row_valid = True
        i = 3

        while (row_valid):
            if not worksheet['A{}'.format(i)].value:
                row_valid = False
                continue

            species = datatabs.cover.CoverSpecies()

            species.micro_plot_id = self.parse_int(worksheet['A{}'.format(i)].value)
            species.quarter = self.parse_int(worksheet['B{}'.format(i)].value)
            species.scale = self.parse_int(worksheet['C{}'.format(i)].value)
            species.species_known = worksheet['D{}'.format(i)].value
            species.species_guess = worksheet['E{}'.format(i)].value

            species.flower = self.parse_float(worksheet['I{}'.format(i)].value)
            species.number_of_stems = self.parse_int(worksheet['J{}'.format(i)].value)

            if species.species_known in datatabs.cover.CoverSpecies.DEER_INDICATOR_SPECIES:
                if None == species.flower:
                    species.flower = 0

            species.percent_cover = self.parse_int(worksheet['F{}'.format(i)].value)
            species.average_height = self.parse_int(worksheet['G{}'.format(i)].value)
            species.count = self.parse_int(worksheet['H{}'.format(i)].value)

            if species.count and (not species.flower) and species.override_species(species.species_known) in datatabs.cover.CoverSpecies.DEER_INDICATOR_SPECIES:
                species.flower = 0

            tab.cover_species.append(species)

            i += 1

        return tab


    def get_collected_cover_subplots(self, workbook):
        subplots = []

        worksheet = workbook[datasheet.TAB_NAME_COVER_TABLE]

        row_valid = True
        i = 3

        while (row_valid):
            if not worksheet['A{}'.format(i)].value:
                row_valid = False
                continue

            subplots.append(self.parse_int(worksheet['A{}'.format(i)].value))

            i = i + 1

        return list(set(subplots))


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
            if not worksheet['A{}'.format(i)].value:
                row_valid = False
                continue

            species = datatabs.sapling.SaplingSpecies()

            species.micro_plot_id = self.parse_int(worksheet['A{}'.format(i)].value)

            if species.micro_plot_id not in subplot_sapling_numbers:
                subplot_sapling_numbers[species.micro_plot_id] = 1
            else:
                subplot_sapling_numbers[species.micro_plot_id] += 1

            species.sapling_number = subplot_sapling_numbers[species.micro_plot_id]
            species.quarter = self.parse_int(worksheet['B{}'.format(i)].value)
            species.scale = self.parse_int(worksheet['C{}'.format(i)].value)
            species.species_known = worksheet['D{}'.format(i)].value
            species.species_guess = worksheet['E{}'.format(i)].value
            species.diameter_breast_height = self.parse_float(worksheet['F{}'.format(i)].value)

            tab.sapling_species.append(species)

            i += 1

        return tab


    def parse_seedling_tab(self, workbook, sheet):
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

            species.quarter = self.parse_int(worksheet['B{}'.format(i)].value)
            species.scale = self.parse_int(worksheet['C{}'.format(i)].value)
            species.species_known = worksheet['D{}'.format(i)].value
            species.species_guess = worksheet['E{}'.format(i)].value
            species.sprout = self.parse_int(worksheet[f'F{i}'].value)

            if None == species.sprout:
                species.sprout = 0

            species.zero_six_inches = self.parse_int(worksheet['G{}'.format(i)].value)
            species.six_twelve_inches = self.parse_int(worksheet['H{}'.format(i)].value)
            species.one_three_feet_total = self.parse_int(worksheet['I{}'.format(i)].value)
            species.one_three_feet_browsed = self.parse_int(worksheet['J{}'.format(i)].value)

            if species.one_three_feet_total and not species.one_three_feet_browsed:
                species.one_three_feet_browsed = 0

            species.three_five_feet_total = self.parse_int(worksheet['K{}'.format(i)].value)
            species.three_five_feet_browsed = self.parse_int(worksheet['L{}'.format(i)].value)

            if species.three_five_feet_total and not species.three_five_feet_browsed:
                species.three_five_feet_browsed = 0

            species.greater_five_feet_total = self.parse_int(worksheet['M{}'.format(i)].value)
            species.greater_five_feet_browsed = self.parse_int(worksheet['N{}'.format(i)].value)

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
            species.diameter_breast_height = self.parse_float(worksheet['E{}'.format(i)].value)

            live_or_dead = self.parse_int(worksheet['F{}'.format(i)].value)

            if None != live_or_dead:
                species.live_or_dead = 'L' if 1 == live_or_dead else 'D'

            species.comments = ''

            tab.tree_species.append(species)

            i += 1

        return tab

