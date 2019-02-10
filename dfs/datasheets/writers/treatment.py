from openpyxl import Workbook
from dfs.datasheets.writers.plot import PlotDatasheetWriter
import dfs.datasheets.datasheet as datasheet
import re


class TreatmentDatasheetWriter(PlotDatasheetWriter):
    def format_general_tab(self, sheet, tab):
        tab['A1'] = 'Study Area'
        tab['B1'] = int(sheet.tabs[datasheet.TAB_NAME_GENERAL].study_area)

        tab['A2'] = 'Plot Number'
        tab['B2'] = int(sheet.tabs[datasheet.TAB_NAME_GENERAL].plot_number)

        tab['A3'] = 'Deer Impact'
        tab['B3'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].deer_impact

        tab['A4'] = 'Date'
        tab['B4'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].collection_date
        tab['B4'].number_format = 'M/D/YYYY'

        tab['A6'] = 'Coordinate Converter'

        tab['A7'] = 'INPUT LAT HERE'
        tab['B7'] = 'INPUT LONG HERE'
        tab['D7'] = 'Yes/No'
        tab['E7'] = 'Yes/No'
        tab['F7'] = 'Yes/No'
        tab['G7'] = '0-1'
        tab['H7'] = '0-4'
        tab['I7'] = 'Meters'
        tab['K7'] = 'DO NOT INPUT VALUES HERE'

        tab['A8'] = 'ex. 4045.1291'
        tab['B8'] = 'ex. 7743.2763'
        tab['C8'] = 'Sub/ Micro'
        tab['D8'] = 'Collected'
        tab['E8'] = 'Re-Monumented'
        tab['F8'] = 'Forested'
        tab['G8'] = 'Disturbance'
        tab['H8'] = 'Type'
        tab['I8'] = 'Altitude'
        tab['J8'] = 'Notes'
        tab['K8'] = 'Latitude'
        tab['L8'] = 'Longitude'

        i = 0
        for rownumber in range(9, 24):
            if i >= len(sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots):
                continue

            subplot = sheet.tabs[datasheet.TAB_NAME_GENERAL].subplots[i]

            tab[f'A{rownumber}'] = subplot.latitude
            tab[f'B{rownumber}'] = subplot.longitude

            tab[f'C{rownumber}'] = int(subplot.micro_plot_id)

            override_collected = 'Yes'

            if subplot.micro_plot_id in sheet.tabs[datasheet.TAB_NAME_COVER_TABLE].get_recorded_subplots():
                override_collected = 'Yes'
            else:
                override_collected = 'No'

            if override_collected != subplot.collected:
                print(f'Warning: value of collected at subplot {subplot.micro_plot_id} does not match the actual collection data and will be overridden.')

            tab[f'D{rownumber}'] = override_collected
            tab[f'E{rownumber}'] = subplot.re_monumented
            tab[f'F{rownumber}'] = subplot.forested
            tab[f'G{rownumber}'] = subplot.disturbance
            tab[f'H{rownumber}'] = subplot.disturbance_type
            tab[f'I{rownumber}'] = subplot.altitude
            tab[f'J{rownumber}'] = subplot.notes

            if subplot.converted_latitude:
                tab[f'K{rownumber}'] = subplot.converted_latitude
            else:
                tab[f'K{rownumber}'] = '=IF(ISBLANK(A{0}),"",LEFT((LEFT(A{0},2)+(RIGHT(A{0},LEN(A{0})-2)/60)),10))'.format(rownumber)

            if subplot.converted_longitude:
                tab[f'L{rownumber}'] = subplot.converted_longitude
            else:
                tab[f'L{rownumber}'] = '=IF(ISBLANK(B{0}),"",LEFT(-1*(LEFT(B{0},2)+(RIGHT(B{0},LEN(B{0})-2)/60)),10))'.format(rownumber)

            i += 1


        tab['A26'] = 'Auxillary Post Locations'
        tab['A27'] = '1-5'
        tab['B27'] = '1 or 2'
        tab['C27'] = 'Wooden or Rebar'
        tab['D27'] = '0-359'
        tab['E27'] = 'Feet'

        tab['A28'] = 'Sub/Micro'
        tab['B28'] = 'Post'
        tab['C28'] = 'Stake Type'
        tab['D28'] = 'Azimuth'
        tab['E28'] = 'Distance'

        i = 0
        for rownumber in range(29, 34):
            if i < len(sheet.tabs[datasheet.TAB_NAME_GENERAL].auxillary_post_locations):
                auxillary_post_location = sheet.tabs[datasheet.TAB_NAME_GENERAL].auxillary_post_locations[i]

                tab[f'A{rownumber}'] = auxillary_post_location.micro_plot_id
                tab[f'B{rownumber}'] = auxillary_post_location.post
                tab[f'C{rownumber}'] = auxillary_post_location.stake_type
                tab[f'D{rownumber}'] = auxillary_post_location.azimuth
                tab[f'E{rownumber}'] = auxillary_post_location.distance
            else:
                tab[f'A{rownumber}'] = ''
                tab[f'B{rownumber}'] = ''
                tab[f'C{rownumber}'] = ''
                tab[f'D{rownumber}'] = ''
                tab[f'E{rownumber}'] = ''

            i += 1

        tab['A35'] = 'Non-Forest Azimuths'
        tab['C35'] = '(if a subplot is not 100% forested, record 3 azimuths along the forested barrier from subplot center)'

        tab['A36'] = 'Subplot'
        tab['B36'] = 'Azimuth 1'
        tab['C36'] = 'Azimuth 2'
        tab['D36'] = 'Azimuth 3'

        i = 0
        for rownumber in range(37, 42):
            if i < len(sheet.tabs[datasheet.TAB_NAME_GENERAL].non_forested_azimuths):
                non_forested_azimuth = sheet.tabs[datasheet.TAB_NAME_GENERAL].non_forested_azimuths[i]

                tab[f'A{rownumber}'] = non_forested_azimuth.micro_plot_id
                tab[f'B{rownumber}'] = non_forested_azimuth.azimuth_1
                tab[f'C{rownumber}'] = non_forested_azimuth.azimuth_2
                tab[f'D{rownumber}'] = non_forested_azimuth.azimuth_3
            else:
                tab[f'A{rownumber}'] = ''
                tab[f'B{rownumber}'] = ''
                tab[f'C{rownumber}'] = ''
                tab[f'D{rownumber}'] = ''


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

        i = 0

        for rownumber in range(3, 33):
            if i < len(sheet.tabs[datasheet.TAB_NAME_WITNESS_TREES].witness_trees):
                tree = sheet.tabs[datasheet.TAB_NAME_WITNESS_TREES].witness_trees[i]

                tab[f'A{rownumber}'] = tree.tree_number
                tab[f'B{rownumber}'] = tree.micro_plot_id
                tab[f'C{rownumber}'] = tree.species_known
                tab[f'D{rownumber}'] = tree.species_guess
                tab[f'E{rownumber}'] = tree.dbh
                tab[f'F{rownumber}'] = tree.live_or_dead
                tab[f'G{rownumber}'] = tree.azimuth
                tab[f'H{rownumber}'] = tree.distance
            else:
                tab[f'A{rownumber}'] = ''
                tab[f'B{rownumber}'] = ''
                tab[f'C{rownumber}'] = ''
                tab[f'D{rownumber}'] = ''
                tab[f'E{rownumber}'] = ''
                tab[f'F{rownumber}'] = ''
                tab[f'G{rownumber}'] = ''
                tab[f'H{rownumber}'] = ''

            i += 1