from dfs.datasheets.parsers.treatments.parser_2015 import TreatmentDatasheetParser2015
import dfs.datasheets.datatabs as datatabs
import dfs.datasheets.datasheet as datasheet


class TreatmentDatasheetParser2016(TreatmentDatasheetParser2015):
    def format_output_filename(self, input_filename):
        return input_filename.replace('Edited', 'Edited_Converted')


    def parse_plot_general_tab(self, workbook, sheet):
        worksheet = workbook[datasheet.TAB_NAME_GENERAL]

        tab = datatabs.general.GeneralTab()

        tab.study_area = self.parse_int(worksheet['B1'].value)
        tab.plot_number = self.parse_int(worksheet['B2'].value)
        tab.deer_impact = self.parse_int(worksheet['B3'].value)
        tab.collection_date = worksheet['B4'].value

        collected_cover_subplots = self.get_collected_cover_subplots(workbook)

        for rownumber in range(8, 23):
            subplot = datatabs.general.TreatmentPlotGeneralPlotSubplot()
            subplot.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

            subplot.latitude = self.parse_float(worksheet[f'I{rownumber}'].value)
            subplot.longitude = self.parse_float(worksheet[f'J{rownumber}'].value)

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

            subplot.altitude = self.parse_float(worksheet[f'H{rownumber}'].value)

            tab.subplots.append(subplot)

        for rownumber in range(34, 44):
            if None != worksheet[f'C{rownumber}'].value:
                auxillary_post_location = datatabs.general.AuxillaryPostLocation()

                auxillary_post_location.post = self.parse_int(worksheet[f'A{rownumber}'].value.replace('Post ', ''))
                auxillary_post_location.micro_plot_id = self.parse_int(worksheet[f'B{rownumber}'].value)
                auxillary_post_location.stake_type = worksheet[f'C{rownumber}'].value
                auxillary_post_location.azimuth = self.parse_int(worksheet[f'D{rownumber}'].value)
                auxillary_post_location.distance = self.parse_float(worksheet[f'E{rownumber}'].value)

                tab.auxillary_post_locations.append(auxillary_post_location)

        for rownumber in range(48, 53):
            if None != self.parse_int(worksheet[f'A{rownumber}'].value):
                non_forested_azimuth = datatabs.general.NonForestedAzimuths()

                non_forested_azimuth.micro_plot_id = self.parse_int(worksheet[f'A{rownumber}'].value)

                non_forested_azimuth.azimuth_1 = self.parse_int(worksheet[f'B{rownumber}'].value)
                non_forested_azimuth.azimuth_2 = self.parse_int(worksheet[f'C{rownumber}'].value)
                non_forested_azimuth.azimuth_3 = self.parse_int(worksheet[f'D{rownumber}'].value)

        return tab

