from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Species
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError
from dfs.datasheets.datatabs.tabs import MissingSubplotValidationError
from dfs.datasheets.datatabs.tabs import FatalExceptionError
import numbers


class SaplingTab(Tab):
    def __init__(self):
        super().__init__()

        self.sapling_species = []


    def postprocess(self):
        pass


    def validate(self):
        validation_errors = []

        for species in self.sapling_species:
            validation_errors += species.validate_attribute_type(['diameter_breast_height'], numbers.Number)

            try:
                validation_errors += species.validate()
            except Exception as e:
                validation_errors.append(FatalExceptionError(self.get_object_type(), e))

        return validation_errors


class SaplingSpecies(Species):
    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.sapling_number = None
        self.quarter = None
        self.scale = None
        self.diameter_breast_height = None


    def validate(self):
        validation_errors = []
        self.species_known = self.override_species(self.species_known)
        self.species_guess = self.override_species(self.species_guess)

        sapling_species = Tab.TREE_SPECIES.copy()
        sapling_species.remove('snag')

        if self.get_species_known() not in sapling_species:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species known', 'sapling species', self.get_species_known()))

        if None != self.get_species_guess() and self.get_species_guess() not in sapling_species:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species guess', 'sapling species', self.get_species_known()))

        if self.micro_plot_id not in range(1, Validatable.MAX_MICRO_PLOT_ID + 1):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'microplot ID', f'1-{Validatable.MAX_MICRO_PLOT_ID}', self.micro_plot_id))

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

        if None == self.diameter_breast_height:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'dbh', 'non-empty', self.diameter_breast_height))
        else:
            if (1 > self.diameter_breast_height) or (5 <= self.diameter_breast_height):
                validation_errors.append(FieldValidationError(self.get_object_type(), 'dbh', '>= 1 or < 5', self.diameter_breast_height))

            if not (self.diameter_breast_height * 10).is_integer():
                validation_errors.append(FieldValidationError(self.get_object_type(), 'dbh', '0.1 increments', self.diameter_breast_height))

        return validation_errors