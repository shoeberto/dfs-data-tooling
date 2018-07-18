from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Species
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError
from dfs.datasheets.datatabs.tabs import MissingSubplotValidationError
from dfs.datasheets.datatabs.tabs import FatalExceptionError
import numbers


class SeedlingTable(Tab):
    def __init__(self):
        super().__init__()

        self.seedling_species = []


    def postprocess(self):
        self.check_primary_key_uniquness(self.seedling_species, ['micro_plot_id', 'quarter', 'scale', 'get_species_known', 'get_species_guess'])


    def validate(self):
        validation_errors = []

        collected_subplots = []
        for species in self.seedling_species:
            validation_errors += species.validate_attribute_type([
                    'zero_six_inches', 
                    'six_twelve_inches', 
                    'one_three_feet_total', 
                    'one_three_feet_browsed', 
                    'three_five_feet_total', 
                    'three_five_feet_browsed', 
                    'greater_five_feet_total', 
                    'greater_five_feet_browsed'
                ], 
                numbers.Number
            )

            try:
                validation_errors += species.validate()
            except Exception as e:
                validation_errors.append(FatalExceptionError(self.get_object_type(), e))

            collected_subplots.append(species.micro_plot_id)

        if sorted(set(collected_subplots)) != sorted(set(range(1, 6))):
            validation_errors.append(MissingSubplotValidationError(self.get_object_type(), collected_subplots))

        return validation_errors


class SeedlingSpecies(Species):
    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.quarter = None
        self.scale = None
        self.sprout = None
        self.zero_six_inches = None
        self.six_twelve_inches = None
        self.one_three_feet_total = None
        self.one_three_feet_browsed = None
        self.three_five_feet_total = None
        self.three_five_feet_browsed = None
        self.greater_five_feet_total = None
        self.greater_five_feet_browsed = None


    def validate(self):
        validation_errors = []

        self.species_guess = self.override_species(self.species_guess)
        self.species_known = self.override_species(self.species_known)

        if self.get_species_known() in Tab.COVER_SPECIES:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species known', 'seedling species', self.get_species_known()))

        if self.get_species_guess() in Tab.COVER_SPECIES:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species guess', 'seedling species', self.get_species_known()))

        if self.micro_plot_id not in range(1, 6):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'microplot ID', '1-5', self.micro_plot_id))

        if self.quarter not in range(1, 5):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'quarter', '1-4', self.quarter))

        self.scale = self.override_scale(self.scale)
        if self.scale not in [300, 1000]:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'scale', '300 or 1000', self.scale))

        if None == self.species_known:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species known', 'non-empty', self.species_known))

        if None != self.species_known:
            validation_errors += self.validate_species(self.species_known)

        if None != self.species_guess:
            validation_errors += self.validate_species(self.species_guess)

        if self.sprout not in [0, 1]:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'sprout', '0 or 1', self.sprout))

        if None != self.zero_six_inches and 0 > self.zero_six_inches:
            validation_errors.append(FieldValidationError(self.get_object_type(), '0-6"', '>= 0', self.zero_six_inches))

        if None != self.six_twelve_inches and 0 > self.six_twelve_inches:
            validation_errors.append(FieldValidationError(self.get_object_type(), '6-12"', '>= 0', self.six_twelve_inches))

        if None != self.one_three_feet_total and 0 > self.one_three_feet_total:
            validation_errors.append(FieldValidationError(self.get_object_type(), "1-3' Total", '>= 0', self.one_three_feet_total))

        if None != self.one_three_feet_browsed and 0 > self.one_three_feet_browsed:
            validation_errors.append(FieldValidationError(self.get_object_type(), "1-3' Browsed", '>= 0', self.one_three_feet_browsed))

        if None != self.one_three_feet_browsed and None != self.one_three_feet_total:
            if self.one_three_feet_browsed > self.one_three_feet_total:
                validation_errors.append(FieldValidationError(self.get_object_type(), "1-3' Browsed'", "<= 1-3' Total", self.one_three_feet_browsed))

        if None != self.three_five_feet_total and 0 > self.three_five_feet_total:
            validation_errors.append(FieldValidationError(self.get_object_type(), "3-5' Total", '>= 0', self.three_five_feet_total))

        if None != self.three_five_feet_browsed and 0 > self.three_five_feet_browsed:
            validation_errors.append(FieldValidationError(self.get_object_type(), "3-5' Browsed", '>= 0', self.three_five_feet_browsed))

        if None != self.three_five_feet_browsed and None != self.three_five_feet_total:
            if self.three_five_feet_browsed > self.three_five_feet_total:
                validation_errors.append(FieldValidationError(self.get_object_type(), "3-5' Browsed", "<= 3-5' Total", self.one_three_feet_browsed))

        if None != self.greater_five_feet_total and 0 > self.greater_five_feet_total:
            validation_errors.append(FieldValidationError(self.get_object_type(), ">5' Total", '>= 0', self.greater_five_feet_total))

        if None != self.greater_five_feet_browsed and 0 > self.greater_five_feet_browsed:
            validation_errors.append(FieldValidationError(self.get_object_type(), ">5' Browsed", '>= 0', self.greater_five_feet_browsed))

        if None != self.greater_five_feet_browsed and None != self.greater_five_feet_total:
            if self.greater_five_feet_browsed > self.greater_five_feet_total:
                validation_errors.append(FieldValidationError(self.get_object_type(), ">5' Browsed", "<= >5' Total", self.greater_five_feet_browsed))

        return validation_errors