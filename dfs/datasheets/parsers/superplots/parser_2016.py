from dfs.datasheets.parsers.plots.parser_2016 import DatasheetParser2016
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class SuperplotDatasheetParser2016(DatasheetParser2016):
    def format_output_filename(self, input_filename):
        return input_filename.replace('QC', 'QC_Converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]

        tab = datatabs.general.SuperplotGeneralTab()

        tab.study_area = self.parse_int(worksheet['B1'].value)
        tab.plot_number = self.parse_int(worksheet['B2'].value)
        tab.deer_impact = self.parse_int(worksheet['B3'].value)
        tab.collection_date = worksheet['B4'].value

        collected_cover_subplots = self.get_collected_cover_subplots(workbook)

        for rownumber in range(8, 19):
            subplot = datatabs.general.PlotGeneralTabSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

            # Ignore slope
            subplot.latitude = self.parse_float(worksheet[f'L{rownumber}'].value)
            subplot.longitude = self.parse_float(worksheet[f'M{rownumber}'].value)

            subplot.forested = str(worksheet[f'D{rownumber}'].value).title().strip()

            if None == subplot.forested:
                if subplot.micro_plot_id in collected_cover_subplots or (None != subplot.latitude or None != subplot.longitude):
                    subplot.forested = 'Yes'
                else:
                    subplot.forested = 'No'

            subplot.disturbance = 1 if 'yes' == str(worksheet[f'E{rownumber}'].value).lower() else 0

            subplot.disturbance_type = self.parse_int(worksheet[f'F{rownumber}'].value)

            if None == subplot.disturbance_type:
                subplot.disturbance_type = 0

            subplot.collected = str(worksheet[f'G{rownumber}'].value).title().strip()

            if None == subplot.collected:
                if subplot.micro_plot_id in collected_cover_subplots:
                    subplot.collected = 'Yes'
                else:
                    subplot.collected = 'No'

            subplot.fenced = str(worksheet[f'H{rownumber}'].value).title().strip()

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

            subplot.altitude = self.parse_float(worksheet[f'K{rownumber}'].value)

            tab.subplots.append(subplot)


        for rownumber in range(22, 26):
            if worksheet[f'B{rownumber}'].value:
                fenced_subplot_condition = datatabs.general.SuperplotFencedSubplotConditions()
                fenced_subplot_condition.repairs = str(worksheet[f'A{rownumber}'].value).title().strip()
                fenced_subplot_condition.level = self.parse_int(worksheet[f'B{rownumber}'].value)
                fenced_subplot_condition.micro_plot_id = worksheet[f'C{rownumber}'].value
                fenced_subplot_condition.notes = worksheet[f'D{rownumber}'].value

                tab.fenced_subplot_condition.append(fenced_subplot_condition)


        for rownumber in range(49, 54):
            if None != self.parse_int(worksheet[f'A{rownumber}'].value):
                non_forested_azimuth = datatabs.general.NonForestedAzimuths()

                non_forested_azimuth.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

                non_forested_azimuth.azimuth_1 = self.parse_int(worksheet[f'B{rownumber}'].value)
                non_forested_azimuth.azimuth_2 = self.parse_int(worksheet[f'C{rownumber}'].value)
                non_forested_azimuth.azimuth_3 = self.parse_int(worksheet[f'D{rownumber}'].value)


        for rownumber in range(35, 45):
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

        for rownumber in range(28, 31):
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