from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError
from dfs.datasheets.datatabs.tabs import MissingSubplotValidationError


class GeneralTab(Tab):
    def __init__(self):
        super().__init__()

        self.study_area = None
        self.plot_number = None
        self.deer_impact = None
        self.collection_date = None
        self.subplots = []
        self.auxillary_post_locations = []
        self.non_forested_azimuths = []


    def postprocess(self):
        pass


    def validate(self):
        validation_errors = []

        if self.study_area not in range(1, 5):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'study area', '1-4', self.study_area))

        if self.plot_number not in range(1, 51):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'plot number', '1-50', self.plot_number))

        if self.deer_impact not in range(0, 6):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'deer impact', '0-5', self.deer_impact))

        if None == self.collection_date:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'collection date', 'a date', 'None'))

        if not self.subplots:
            validation_errors.append(FieldCountValidationError('General Tab', 'subplots', 5, 'None'))
        
        if 5 != len(self.subplots):
            validation_errors.append(FieldCountValidationError('General Tab', 'subplots', 5, len(self.subplots)))

        micro_plot_id_list = []
        for subplot in self.subplots:
            validation_errors += subplot.validate()
            micro_plot_id_list.append(subplot.micro_plot_id)

        if list(set(micro_plot_id_list)) != range(1, 6):
            validation_errors.append(MissingSubplotValidationError(self.get_object_type(), micro_plot_id_list))

        last_post_number = None
        for auxillary_post_location in self.auxillary_post_locations:
            validation_errors += auxillary_post_location.validate()

            if None == last_post_number and 1 != auxillary_post_location.post:
                validation_errors.append(FieldValidationError('Auxillary Post Location', 'Post number', 1, auxillary_post_location.post))
            elif 1 == last_post_number and 2 != auxillary_post_location.post:
                validation_errors.append(FieldValidationError('Auxillary Post Location', 'Post number', 2, auxillary_post_location.post))
            elif 2 == last_post_number and 1 != auxillary_post_location.post:
                validation_errors.append(FieldValidationError('Auxillary Post Location', 'Post number', 1, auxillary_post_location.post))

            last_post_number = auxillary_post_location.post

        for non_forested_azimuth in self.non_forested_azimuths:
            validation_errors += non_forested_azimuth.validate()

        return validation_errors


class PlotGeneralTab(GeneralTab):
    def __init__(self):
        super().__init__()

        fenced_subplot_condition = None

    def validate(self):
        # TODO: validation code
        return super().validate()


class GeneralTabSubplot(Validatable):
    def __init__(self):
        super().__init__()

        self.latitude = None
        self.longitude = None
        self.micro_plot_id = None    # 1-5
        self.collected = None        # Yes/No
        self.fenced = None           # Yes/No
        self.azimuth = None          # 0-365
        self.distance = None         # Feet
        self.altitude = None         # Meters
        self.notes = None


    def validate(self):
        validation_errors = []

        if self.micro_plot_id not in range(1, 6):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'microplot ID', '1-5', self.micro_plot_id))

        validation_errors += self.validate_mutually_required_fields(['collected', 'fenced', 'azimuth', 'distance'])

        if None != self.collected and self.collected not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'collected', self.YES_NO_TEXT, self.collected))

        if None != self.fenced and self.fenced not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'fenced', self.YES_NO_TEXT, self.fenced))

        if None != self.azimuth and (self.azimuth < 0 or self.azimuth > 365):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'azimuth', '0-365', self.azimuth))

        if None != self.distance and self.distance <= 0:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'distance', '> 0 feet', self.distance))

        return validation_errors


class PlotGeneralTabSubplot(GeneralTabSubplot):
    def __init__(self):
        super().__init__()

        self.re_monumented = None    # Yes/No
        self.forested = None         # Yes/No
        self.disturbance = None      # 0-1
        self.disturbance_type = None # 0-4; if disturbance = 0, must be 0; if disturbance = 1, must be 1-4
        self.lime = None             # Yes/No
        self.herbicide = None        # Yes/No


    def validate(self):
        validation_errors = []

        validation_errors += super().validate()

        validation_errors += self.validate_mutually_required_fields(['re_monumented', 'disturbance', 'disturbance_type', 'lime', 'herbicide'])

        if None != self.re_monumented and self.re_monumented not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'remonumented', self.YES_NO_TEXT, self.re_monumented))

        if self.forested not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'forested', self.YES_NO_TEXT, self.forested))

        if None != self.disturbance and self.disturbance not in [0, 1]:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'disturbance', '0-1', self.disturbance))

        if None != self.disturbance_type:
            if None == self.disturbance_type or \
                (1 == self.disturbance and self.disturbance_type not in range(1, 5)) or \
                (0 == self.disturbance and 0 != self.disturbance_type):
                validation_errors.append(FieldValidationError(self.get_object_type(), 'disturbance type', '0-4', self.disturbance_type))

        if None != self.lime and self.lime not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'lime', self.YES_NO_TEXT, self.lime))

        if None != self.herbicide and self.herbicide not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'herbicide', self.YES_NO_TEXT, self.herbicide))

        return validation_errors



class TreatmentPlotGeneralPlotSubplot(GeneralTabSubplot):
    def __init__(self):
        # passthrough
        super().__init__()


    def validate(self):
        # TODO: validation code
        return super().validate()


class FencedSubplotConditions(Validatable):
    def __init__(self):
        super().__init__()

        self.active_exclosure = None # yes/no
        self.repairs = None          # yes/no
        self.level = None            # 0-3
        self.notes = None            # text


    def validate(self):
        validation_errors = []
        validation_errors += self.validate_mutually_required_fields(['active_exclosure', 'repairs', 'level', 'notes'])

        if None != self.active_exclosure and self.active_exclosure not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'fenced subplot condition', self.YES_NO_TEXT, self.active_exclosure))

        if None != self.repairs and self.repairs not in self.YES_NO_RESPONSE:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'repairs', self.YES_NO_TEXT, self.repairs))

        if None != self.level and self.level not in range(0, 4):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'level', '0-3', self.level))

        return validation_errors



class AuxillaryPostLocation(Validatable):
    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.post = None             # 1 or 2
        self.stake_type = None       # "Wooden" or "Rebar"
        self.azimuth = None          # 0 - 359
        self.distance = None         # Feet


    def validate(self):
        validation_errors = []
        validation_errors += self.validate_mutually_required_fields(['micro_plot_id', 'post', 'stake_type', 'azimuth', 'distance'])

        if None != self.micro_plot_id and self.micro_plot_id not in range(1, 6):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'microplot ID', '1-5', self.micro_plot_id))

        if None != self.post and self.post not in [1, 2]:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'post', '1 or 2', self.post))

        if None != self.stake_type and self.stake_type not in ['Wooden', 'Rebar']:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'stake type', 'Wooden or Rebar', self.stake_type))

        if None != self.azimuth and (self.azimuth < 0 or self.azimuth > 359):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'azimuth', '0-359', self.azimuth))

        if None != self.distance and self.distance <= 0:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'distance', '> 0', self.distance))

        return validation_errors



class NonForestedAzimuths(Validatable):
    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.azimuth_1 = None
        self.azimuth_2 = None
        self.azimuth_3 = None


    def validate(self):
        validation_errors = []
        validation_errors += self.validate_mutually_required_fields(['micro_plot_id', 'azimuth_1', 'azimuth_2', 'azimuth_3'])

        if None != self.micro_plot_id and self.micro_plot_id not in range(1, 6):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'micro plot ID', '1-5', self.micro_plot_id))

        if None != self.azimuth_1 and (self.azimuth_1 < 0 or self.azimuth_1 > 359):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'azimuth 1', '0-359', self.azimuth_1))

        if None != self.azimuth_2 and (self.azimuth_2 < 0 or self.azimuth_2 > 359):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'azimuth 2', '0-359', self.azimuth_2))

        if None != self.azimuth_3 and (self.azimuth_3 < 0 or self.azimuth_3 > 359):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'azimuth 3', '0-359', self.azimuth_3))

        return validation_errors

