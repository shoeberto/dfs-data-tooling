from openpyxl import Workbook
import dfs.datasheets.datasheet as datasheet


class DatasheetWriter:
    def write(self, sheet, output_directory):
        """
        Write a datasheet to an xlsx file.

        Keyword arguments:
        sheet -- a datasheet object.
        output_directory -- the directory to write the new xlsx file to.
        """

        # TODO: make this support treatment datasheets as well as plot datasheets
        workbook = Workbook()

        # remove the default worksheet
        workbook.remove(workbook['Sheet'])

        general_tab = workbook.create_sheet(title=datasheet.TAB_NAME_GENERAL)
        self.format_general_tab(sheet, general_tab)

        witness_tree_tab = workbook.create_sheet(title=datasheet.TAB_NAME_WITNESS_TREES)
        self.format_witness_trees_tab(sheet, witness_tree_tab)

        cover_table_tab = workbook.create_sheet(title=datasheet.TAB_NAME_COVER_TABLE)
        self.format_cover_table_tab(sheet, cover_table_tab)

        workbook.save('{}/{}'.format(output_directory, sheet.input_filename))


    def format_general_tab(self, sheet, tab):
        tab['A1'] = 'Study Area'
        tab['B1'] = str(sheet.tabs[datasheet.TAB_NAME_GENERAL].study_area)

        tab['A2'] = 'Plot Number'
        tab['B2'] = str(sheet.tabs[datasheet.TAB_NAME_GENERAL].plot_number)

        tab['A3'] = 'Deer Impact'
        tab['B3'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].deer_impact

        tab['A4'] = 'Date'
        tab['B4'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].collection_date

        tab['A6'] = 'Coordinate Converter'

        tab['A7'] = 'INPUT LAT HERE'
        tab['B7'] = 'INPUT LONG HERE'
        tab['C7'] = 'Data Type'
        tab['D7'] = 'Yes/No'
        tab['E7'] = 'Yes/No'
        tab['F7'] = '0-359'
        tab['G7'] = 'Feet'
        tab['H7'] = 'Meters'
        tab['I7'] = 'DO NOT INPUT VALUES HERE'

        tab['A8'] = 'ex. 4045.1291'
        tab['B8'] = 'ex. 7743.2763'
        tab['C8'] = 'Sub/ Micro'
        tab['D8'] = 'Collected'
        tab['E8'] = 'Fenced'
        tab['F8'] = 'Azimuth'
        tab['G8'] = 'Distance'
        tab['H8'] = 'Altitude'
        tab['I8'] = 'Latitude'
        tab['J8'] = 'Longitude'
        tab['K8'] = 'UID'

        i = 0
        for rownumber in range(9, 14):
            subplot = sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots[i]

            tab['A{}'.format(rownumber)] = subplot.latitude
            tab['B{}'.format(rownumber)] = subplot.longitude
            tab['C{}'.format(rownumber)] = str(subplot.micro_plot_id)
            tab['D{}'.format(rownumber)] = subplot.collected
            tab['E{}'.format(rownumber)] = subplot.fenced
            tab['F{}'.format(rownumber)] = subplot.azimuth
            tab['G{}'.format(rownumber)] = subplot.distance
            tab['H{}'.format(rownumber)] = subplot.altitude
            tab['I{}'.format(rownumber)] = '=LEFT((LEFT(A{0},2)+(RIGHT(A{0},LEN(A{0})-2)/60)),10)'.format(rownumber)
            tab['J{}'.format(rownumber)] = '=LEFT(-1*(LEFT(B{0},2)+(RIGHT(B{0},LEN(B{0})-2)/60)),10)'.format(rownumber)
            tab['K{}'.format(rownumber)] = '=VALUE(CONCATENATE(General!$B$1,IF(LEN(General!$B$2)<2, CONCATENATE(0,General!$B$2),General!$B$2),IF(LEN(C{0})<2, CONCATENATE(0,C{0}),C{0}))'.format(rownumber)

            i += 1

        tab['A15'] = 'Data Type'
        tab['B15'] = 'Yes/No'
        tab['C15'] = 'Yes/No'
        tab['D15'] = '0-1'
        tab['E15'] = 'Type'
        tab['F15'] = 'Yes/No'
        tab['G15'] = 'Yes/No'

        tab['A16'] = 'Sub/Micro'
        tab['B16'] = 'Re-Monumented'
        tab['C16'] = 'Forested'
        tab['D16'] = 'Disturbance'
        tab['E16'] = 'Type'
        tab['F16'] = 'Lime'
        tab['G16'] = 'Herbicide'

        i = 0
        for rownumber in range(17, 22):
            subplot = sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots[i]

            tab['A{}'.format(rownumber)] = subplot.micro_plot_id
            tab['B{}'.format(rownumber)] = subplot.re_monumented
            tab['C{}'.format(rownumber)] = subplot.forested
            tab['D{}'.format(rownumber)] = subplot.disturbance
            tab['E{}'.format(rownumber)] = subplot.disturbance_type
            tab['F{}'.format(rownumber)] = subplot.lime
            tab['G{}'.format(rownumber)] = subplot.herbicide

            i += 1

        tab['A23'] = 'Fenced Subplot Condition'
        tab['C23'] = '0-3'

        tab['A24'] = 'Active Exclosure'
        tab['B24'] = 'Repair(s)'
        tab['C24'] = 'Level'
        tab['C24'] = 'Level'

        tab['A25'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].fenced_subplot_condition.active_exclosure
        tab['B25'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].fenced_subplot_condition.repairs
        tab['C25'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].fenced_subplot_condition.level

        tab['A27'] = 'Auxillary Post Locations'
        tab['A28'] = '1-5'
        tab['B28'] = '1 or 2'
        tab['C28'] = 'Wooden or Rebar'
        tab['D28'] = '0-359'
        tab['E28'] = 'Feet'

        tab['A29'] = 'Sub/Micro'
        tab['B29'] = 'Post'
        tab['C29'] = 'Stake Type'
        tab['D29'] = 'Azimuth'
        tab['E29'] = 'Distance'

        i = 0
        for rownumber in range(30, 35):
            if i < len(sheet.tabs[datasheet.TAB_NAME_GENERAL].auxillary_post_locations):
                auxillary_post_location = sheet.tabs[datasheet.TAB_NAME_GENERAL].auxillary_post_location[i]

                tab['A{}'.format(rownumber)] = auxillary_post_location.micro_plot_id
                tab['B{}'.format(rownumber)] = auxillary_post_location.post
                tab['C{}'.format(rownumber)] = auxillary_post_location.stake_type
                tab['D{}'.format(rownumber)] = auxillary_post_location.azimuth
                tab['E{}'.format(rownumber)] = auxillary_post_location.distance

            i += 1


    def format_witness_trees_tab(self, sheet, tab):
        tab['A1'] = 'Witness Tree Table'

        tab['A2'] = 'Tree No'
        tab['B2'] = 'Subplot'
        tab['C2'] = 'Spp_K'
        tab['D2'] = 'Spp_G'
        tab['E2'] = 'dbh'
        tab['F2'] = 'L or D'
        tab['G2'] = 'Azimuth'
        tab['H2'] = 'Distance'
        tab['I2'] = 'UID'

        default_tree_number = 1
        i = 0

        for rownumber in range(3, 14):
            tab['A{}'.format(rownumber)] = default_tree_number

            if rownumber < 5 or (rownumber > 5 and rownumber < 7) or (rownumber > 7 and rownumber < 9) or (rownumber > 9 and rownumber < 11) or (rownumber > 11 and rownumber < 13):
                default_tree_number += 1
            else:
                default_tree_number = 1

            if i < len(sheet.tabs[datasheet.TAB_NAME_WITNESS_TREES].witness_trees):
                tree = sheet.tabs[datasheet.TAB_NAME_WITNESS_TREES].witness_trees[i]

                tab['B{}'.format(rownumber)] = tree.micro_plot_id
                tab['C{}'.format(rownumber)] = tree.species_known
                tab['D{}'.format(rownumber)] = tree.species_guess
                tab['E{}'.format(rownumber)] = tree.dbh
                tab['F{}'.format(rownumber)] = tree.live_or_dead
                tab['G{}'.format(rownumber)] = tree.azimuth
                tab['H{}'.format(rownumber)] = tree.distance
                tab['I{}'.format(rownumber)] = '=VALUE(CONCATENATE(General!$B$1,IF(LEN(General!$B$2)<2, CONCATENATE(0,General!$B$2),General!$B$2),IF(LEN(B3)<2, CONCATENATE(0,B3),B3)))'

            i += 1


    def format_cover_table_tab(self, sheet, tab):
        tab['A1'] = 'Cover Table'

        tab['A2'] = 'Micro'
        tab['B2'] = 'Quarter'
        tab['C2'] = '300/1000'
        tab['D2'] = 'Spp_K'
        tab['E2'] = 'Spp_G'
        tab['F2'] = 'PctCov'
        tab['G2'] = 'AvgHt'
        tab['H2'] = 'Count'
        tab['I2'] = 'Flower'
        tab['J2'] = 'Nstem'
        tab['K2'] = 'UID'

        i = 3
        for cover_species in sheet.tabs[datasheet.TAB_NAME_COVER_TABLE].cover_species:
            tab['A{}'.format(i)] = cover_species.micro_plot_id
            tab['B{}'.format(i)] = cover_species.quarter
            tab['C{}'.format(i)] = cover_species.scale
            tab['D{}'.format(i)] = cover_species.species_known
            tab['E{}'.format(i)] = cover_species.species_guess
            tab['F{}'.format(i)] = cover_species.percent_cover
            tab['G{}'.format(i)] = cover_species.average_height
            tab['H{}'.format(i)] = cover_species.count
            tab['I{}'.format(i)] = cover_species.flower
            tab['J{}'.format(i)] = cover_species.number_of_stems
            tab['K{}'.format(i)] = '=VALUE(CONCATENATE(General!$B$1,IF(LEN(General!$B$2)<2, CONCATENATE(0,General!$B$2),General!$B$2),IF(LEN(A{0})<2, CONCATENATE(0,A{0}),A{0})))'.format(i)

            i += 1