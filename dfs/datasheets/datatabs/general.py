from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationException
from dfs.datasheets.datatabs.tabs import FieldCountValidationException


class GeneralTab(Tab):
    study_area = None
    plot_number = None
    deer_impact = None
    collection_date = None
    subplots = []
    auxillary_post_locations = []
    non_forested_azimuths = []


    def validate(self):
        if self.study_area not in [1, 2]:
            raise FieldValidationException(self.__name__, 'study area', '1 or 2', self.study_area)

        if self.plot_number not in range(1, 51):
            raise FieldValidationException(self.__name__, 'plot number', '1-50', self.plot_number)

        # TODO: what's a valid value here?
        # if self.deer_impact not in []:
        #     raise FieldValidationException(self.__name__, 'deer impact', 'TODO', self.deer_impact)

        if None == self.collection_date:
            raise FieldValidationException(self.__name__, 'collection date', 'a date', 'None')

        if not self.subplots:
            raise FieldCountValidationException('General Tab', 'subplots', 5, 'None')
        
        if 5 != len(self.subplots):
            raise FieldCountValidationException('General Tab', 'subplots', 5, len(self.subplots))

        for subplot in self.subplots:
            subplot.validate()

        last_post_number = None
        for auxillary_post_location in self.auxillary_post_locations:
            auxillary_post_location.validate()

            if None == last_post_number and 1 != auxillary_post_location.post:
                raise FieldValidationException('Auxillary Post Location', 'Post number', 1, auxillary_post_location.post)
            elif 1 == last_post_number and 2 != auxillary_post_location.post:
                raise FieldValidationException('Auxillary Post Location', 'Post number', 2, auxillary_post_location.post)
            elif 2 == last_post_number and 1 != auxillary_post_location.post:
                raise FieldValidationException('Auxillary Post Location', 'Post number', 1, auxillary_post_location.post)

            last_number = auxillary_post_location.post


        for non_forested_azimuth in self.non_forested_azimuths:
            non_forested_azimuth.validate()


class PlotGeneralTab(GeneralTab):
    fenced_subplot_condition = None

    def validate(self):
        # TODO: validation code
        super().validate()


class GeneralTabSubplot(Validatable):
    latitude = None
    longitude = None
    micro_plot_id = None    # 1-5
    collected = None        # Yes/No
    fenced = None           # Yes/No
    azimuth = None          # 0-365
    distance = None         # Feet
    altitude = None         # Meters
    notes = None


    def validate(self):
        # TODO: validation code
        return


class PlotGeneralTabSubplot(GeneralTabSubplot):
    re_monumented = None    # Yes/No
    forested = None         # Yes/No
    disturbance = None      # 0-1
    disturbance_type = None # 0-4; if disturbance = 0, must be 0; if disturbance = 1, must be 1-4
    lime = None             # Yes/No
    herbicide = None        # Yes/No


    def validate(self):
        # TODO: validation code
        super().validate()


class TreatmentPlotGeneralPlotSubplot(GeneralTabSubplot):
    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


    def validate(self):
        # TODO: validation code
        super().validate()


class FencedSubplotConditions(Validatable):
    active_exclosure = None # text?
    repairs = None          # text?
    level = None            # 0-3
    notes = None            # text


    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


    def validate(self):
        # TODO: validation code
        return


class AuxillaryPostLocation(Validatable):
    micro_plot_id = None
    post = None             # 1 or 2
    stake_type = None       # "Wooden" or "Rebar"
    azimuth = None          # 0 - 359
    distance = None         # Feet


    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


    def validate(self):
        # TODO: validation code
        return



class NonForestedAzimuths(Validatable):
    micro_plot_id = None
    azimuth_1 = None
    azimuth_2 = None
    azimuth_3 = None


    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


    def validate(self):
        # TODO: validation code
        return
