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

        workbook = Workbook()

        # remove the default worksheet
        workbook.remove(workbook['Sheet'])

        general_tab = workbook.create_sheet(title=datasheet.TAB_NAME_GENERAL)
        self.format_general_tab(sheet, general_tab)

        witness_tree_tab = workbook.create_sheet(title=datasheet.TAB_NAME_WITNESS_TREES)
        self.format_witness_trees_tab(sheet, witness_tree_tab)

        workbook.save('{}/{}'.format(output_directory, sheet.input_filename))


    def format_general_tab(self, sheet, tab):
        tab['A1'] = 'Study Area'
        tab['B1'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].study_area

        tab['A2'] = 'Plot Number'
        tab['B2'] = sheet.tabs[datasheet.TAB_NAME_GENERAL].plot_number


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

            i += 1