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
            raise FieldValidationException(self.__class__.__name__, 'study area', '1 or 2', self.study_area)

        if self.plot_number not in range(1, 51):
            raise FieldValidationException(self.__class__.__name__, 'plot number', '1-50', self.plot_number)

        if self.deer_impact not in range(1, 6):
            raise FieldValidationException(self.__class__.__name__, 'deer impact', '1-5', self.deer_impact)

        if None == self.collection_date:
            raise FieldValidationException(self.__class__.__name__, 'collection date', 'a date', 'None')

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
        if self.micro_plot_id not in range(1, 6):
            raise FieldValidationException(self.__class__.__name__, 'microplot ID', '1-5', self.micro_plot_id)

        self.validate_mutually_required_fields(['collected', 'fenced', 'azimuth', 'distance'])

        if None != self.collected and self.collected not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'collected', self.YES_NO_TEXT, self.collected)

        if None != self.fenced and self.fenced not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'fenced', self.YES_NO_TEXT, self.fenced)

        if None != self.azimuth and (self.azimuth < 0 or self.azimuth > 365):
            raise FieldValidationException(self.__class__.__name__, 'azimuth', '0-365', self.azimuth)

        if None != self.distance and self.distance <= 0:
            raise FieldValidationException(self.__class__.__name__, 'distance', '> 0 feet', self.distance)


class PlotGeneralTabSubplot(GeneralTabSubplot):
    re_monumented = None    # Yes/No
    forested = None         # Yes/No
    disturbance = None      # 0-1
    disturbance_type = None # 0-4; if disturbance = 0, must be 0; if disturbance = 1, must be 1-4
    lime = None             # Yes/No
    herbicide = None        # Yes/No


    def validate(self):
        super().validate()

        self.validate_mutually_required_fields(['re_monumented', 'disturbance', 'disturbance_type', 'lime', 'herbicide'])

        if None != self.re_monumented and self.re_monumented not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'remonumented', self.YES_NO_TEXT, self.re_monumented)

        if self.forested not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'forested', self.YES_NO_TEXT, self.forested)

        if None != self.disturbance and self.disturbance not in [0, 1]:
            raise FieldValidationException(self.__class__.__name__, 'disturbance', '0-1', self.disturbance)

        if None != self.disturbance_type:
            if None == self.disturbance_type or \
                (1 == self.disturbance and self.disturbance_type not in range(1, 5)) or \
                (0 == self.disturbance and 0 != self.disturbance_type):
                raise FieldValidationException(self.__class__.__name__, 'disturbance type', '0-4', self.disturbance_type)

        if None != self.lime and self.lime not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'lime', self.YES_NO_TEXT, self.lime)

        if None != self.herbicide and self.herbicide not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'herbicide', self.YES_NO_TEXT, self.herbicide)



class TreatmentPlotGeneralPlotSubplot(GeneralTabSubplot):
    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


    def validate(self):
        # TODO: validation code
        super().validate()


class FencedSubplotConditions(Validatable):
    active_exclosure = None # yes/no
    repairs = None          # yes/no
    level = None            # 0-3
    notes = None            # text


    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


    def validate(self):
        self.validate_mutually_required_fields(['active_exclosure', 'repairs', 'level', 'notes'])

        if None != self.active_exclosure and self.active_exclosure not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'fenced subplot condition', self.YES_NO_TEXT, self.active_exclosure)

        if None != self.repairs and self.repairs not in self.YES_NO_RESPONSE:
            raise FieldValidationException(self.__class__.__name__, 'repairs', self.YES_NO_TEXT, self.repairs)

        if None != self.level and self.level not in range(0, 4):
            raise FieldValidationException(self.__class__.__name__, 'level', '0-3', self.level)



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
        self.validate_mutually_required_fields(['micro_plot_id', 'post', 'stake_type', 'azimuth', 'distance'])

        if None != self.micro_plot_id and self.micro_plot_id not in range(1, 6):
            raise FieldValidationException(self.__class__.__name__, 'microplot ID', '1-5', self.micro_plot_id)

        if None != self.post and self.post not in [1, 2]:
            raise FieldValidationException(self.__class__.__name__, 'post', '1 or 2', self.post)

        if None != self.stake_type and self.stake_type not in ['Wooden', 'Rebar']:
            raise FieldValidationException(self.__class__.__name__, 'stake type', 'Wooden or Rebar', self.stake_type)

        if None != self.azimuth and (self.azimuth < 0 or self.azimuth > 359):
            raise FieldValidationException(self.__class__.__name__, 'azimuth', '0-359', self.azimuth)

        if None != self.distance and self.distance <= 0:
            raise FieldValidationException(self.__class__.__name__, 'distance', '> 0', self.distance)



class NonForestedAzimuths(Validatable):
    micro_plot_id = None
    azimuth_1 = None
    azimuth_2 = None
    azimuth_3 = None


    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


    def validate(self):
        self.validate_mutually_required_fields(['micro_plot_id', 'azimuth_1', 'azimuth_2', 'azimuth_3'])

        if None != self.micro_plot_id and self.micro_plot_id not in range(1, 6):
            raise FieldValidationException(self.__class__.__name__, 'micro plot ID', '1-5', self.micro_plot_id)

        if None != self.azimuth_1 and (self.azimuth_1 < 0 or self.azimuth_1 > 359):
            raise FieldValidationException(self.__class__.__name__, 'azimuth 1', '0-359', self.azimuth_1)

        if None != self.azimuth_2 and (self.azimuth_2 < 0 or self.azimuth_2 > 359):
            raise FieldValidationException(self.__class__.__name__, 'azimuth 2', '0-359', self.azimuth_2)

        if None != self.azimuth_3 and (self.azimuth_3 < 0 or self.azimuth_3 > 359):
            raise FieldValidationException(self.__class__.__name__, 'azimuth 3', '0-359', self.azimuth_3)

