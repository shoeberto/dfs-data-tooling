from dfs.datasheets.parser import DatasheetParser
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class DatasheetParser2014(DatasheetParser):
    def format_output_filename(self, input_filename):
        return input_filename.replace('Data', 'Data_Converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.general.PlotGeneralTab()

        if worksheet['D1'].value:
            tab.study_area = self.parse_int(worksheet['D1'].value)
            tab.plot_number = self.parse_int(worksheet['D2'].value)
        else:
            study_area, plot_number = str(worksheet['D2'].value).split('_')
            tab.study_area = self.parse_int(study_area)
            tab.plot_number = self.parse_int(plot_number)

        tab.deer_impact = self.parse_int(worksheet['D3'].value)
        tab.collection_date = worksheet['M3'].value

        # not recorded for 2014
        tab.fenced_subplot_condition = datatabs.general.FencedSubplotConditions()

        for rownumber in range(6, 11):
            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'J{rownumber}'].value)

            if None == subplot.micro_plot_id:
                subplot.micro_plot_id = rownumber - 5

            # Ignore slope
            subplot.slope = self.parse_int(worksheet[f'K{rownumber}'].value)
            # forested not collected in 2014

            tab.subplots.append(subplot)

        return tab


    def parse_witness_tree_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]
        tab = datatabs.witnesstree.WitnessTreeTab()

        for rownumber in range(6, 9):
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
        i = 5

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

        collected_plots = list(set([x.micro_plot_id for x in tab.cover_species]))

        for i in range(0, len(sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots)):
            micro_plot_id = sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots[i].micro_plot_id
            sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots[i].collected = 'Yes' if micro_plot_id in collected_plots else 'No'
            sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots[i].forested = 'Yes' if micro_plot_id in collected_plots else 'No'

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

        tab.general_notes = self.parse_note_fields(worksheet, 'C22:M42')

        return tab


    def parse_sapling_tab(self, workbook, sheet):
        worksheet = workbook['Sapling_(1-5")']
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
        worksheet = workbook['Seedling_(0-1")']
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
            if not worksheet[f'B{i}'].value:
                row_valid = False
                continue

            species = datatabs.tree.TreeTableSpecies()

            species.micro_plot_id = worksheet[f'B{i}'].value

            if species.micro_plot_id not in subplot_tree_numbers:
                subplot_tree_numbers[species.micro_plot_id] = 1
            else:
                subplot_tree_numbers[species.micro_plot_id] += 1

            species.tree_number = subplot_tree_numbers[species.micro_plot_id]
            species.species_known = worksheet[f'D{i}'].value
            species.species_guess = worksheet[f'E{i}'].value
            species.diameter_breast_height = self.parse_float(worksheet[f'F{i}'].value)

            live_or_dead = str(worksheet[f'G{i}'].value).upper()

            if None != live_or_dead and live_or_dead.lower() in ['l', 'd']:
                species.live_or_dead = live_or_dead.upper()
            else:
                species.live_or_dead = live_or_dead

            species.comments = ''

            tab.tree_species.append(species)

            i += 1

        return tab


    def validate_workbook(self, workbook, filepath):
        for t in [datasheet.TAB_NAME_GENERAL, datasheet.TAB_NAME_TREE_TABLE, datasheet.TAB_NAME_COVER_TABLE, 'Sapling_(1-5")', 'Seedling_(0-1")']:
            if t not in workbook:
                raise Exception(f"Missing required worksheet '{t}' in file '{filepath}'")