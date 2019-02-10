from openpyxl import Workbook
from dfs.datasheets.writers.plot import PlotDatasheetWriter
import dfs.datasheets.datasheet as datasheet
import re


class SuperplotDatasheetWriter(PlotDatasheetWriter):
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

        tab['A21'] = 'Data Type'
        tab['B21'] = 'Yes/No'
        tab['C21'] = 'Yes/No'
        tab['D21'] = '0-1'
        tab['E21'] = '0-4'
        tab['F21'] = 'Yes/No'
        tab['G21'] = 'Yes/No'
        tab['H21'] = "(notes contained in 'H' cells only)"

        tab['A22'] = 'Sub/Micro'
        tab['B22'] = 'Re-Monumented'
        tab['C22'] = 'Forested'
        tab['D22'] = 'Disturbance'
        tab['E22'] = 'Type'
        tab['F22'] = 'Lime'
        tab['G22'] = 'Herbicide'
        tab['H22'] = 'Notes'

        i = 0
        for rownumber in range(9, 20):
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
            tab[f'E{rownumber}'] = subplot.fenced
            tab[f'F{rownumber}'] = subplot.azimuth
            tab[f'G{rownumber}'] = subplot.distance
            tab[f'H{rownumber}'] = subplot.altitude

            if subplot.converted_latitude:
                tab[f'I{rownumber}'] = subplot.converted_latitude
            else:
                tab[f'I{rownumber}'] = '=IF(ISBLANK(A{0}),"",LEFT((LEFT(A{0},2)+(RIGHT(A{0},LEN(A{0})-2)/60)),10))'.format(rownumber)

            if subplot.converted_longitude:
                tab[f'J{rownumber}'] = subplot.converted_longitude
            else:
                tab[f'J{rownumber}'] = '=IF(ISBLANK(B{0}),"",LEFT(-1*(LEFT(B{0},2)+(RIGHT(B{0},LEN(B{0})-2)/60)),10))'.format(rownumber)

            offset_index = rownumber + 14

            tab[f'A{offset_index}'] = subplot.micro_plot_id
            tab[f'B{offset_index}'] = subplot.re_monumented
            tab[f'C{offset_index}'] = subplot.forested
            tab[f'D{offset_index}'] = subplot.disturbance
            tab[f'E{offset_index}'] = subplot.disturbance_type
            tab[f'F{offset_index}'] = subplot.lime
            tab[f'G{offset_index}'] = subplot.herbicide
            tab[f'H{offset_index}'] = subplot.notes

            i += 1

        tab['A35'] = 'Fenced Subplot Condition'
        tab['C35'] = 'Yes/No'
        tab['D35'] = '0-3'
        tab['E35'] = "(notes contained in 'E' cells only)"

        tab['A36'] = 'Subplot'
        tab['B36'] = 'Active Exclosure'
        tab['C36'] = 'Repair(s)'
        tab['D36'] = 'Level'
        tab['E36'] = 'Notes'

        i = 0
        for rownumber in range(37, 41):
            if i >= len(sheet.tabs[datasheet.TAB_NAME_GENERAL].fenced_subplot_condition):
                continue

            condition = sheet.tabs[datasheet.TAB_NAME_GENERAL].fenced_subplot_condition[i]

            tab[f'A{rownumber}'] = condition.micro_plot_id
            tab[f'B{rownumber}'] = condition.active_exclosure
            tab[f'C{rownumber}'] = condition.repairs
            tab[f'D{rownumber}'] = condition.level
            tab[f'E{rownumber}'] = condition.notes

            i += 1


        tab['A42'] = 'Auxillary Post Locations'
        tab['A43'] = '1-5'
        tab['B43'] = '1 or 2'
        tab['C43'] = 'Wooden or Rebar'
        tab['D43'] = '0-359'
        tab['E43'] = 'Feet'

        tab['A44'] = 'Sub/Micro'
        tab['B44'] = 'Post'
        tab['C44'] = 'Stake Type'
        tab['D44'] = 'Azimuth'
        tab['E44'] = 'Distance'

        i = 0
        for rownumber in range(45, 50):
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

        tab['A51'] = 'Non-Forest Azimuths'
        tab['C51'] = '(if a subplot is not 100% forested, record 3 azimuths along the forested barrier from subplot center)'

        tab['A52'] = 'Subplot'
        tab['B52'] = 'Azimuth 1'
        tab['C52'] = 'Azimuth 2'
        tab['D52'] = 'Azimuth 3'

        i = 0
        for rownumber in range(53, 58):
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