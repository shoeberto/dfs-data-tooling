from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError

class SaplingTab(Tab):
    def __init__(self):
        super().__init__()

        self.sapling_species = []


    def validate(self):
        validation_errors = []

        for species in self.sapling_species:
            validation_errors += species.validate()

        return validation_errors


class SaplingSpecies(Validatable):
    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.sapling_number = None
        self.quarter = None
        self.scale = None
        self.species_known = None
        self.species_guess = None
        self.diameter_breast_height = None


    def validate(self):
        validation_errors = []
        self.species_known = self.override_species(self.species_known)
        self.species_guess = self.override_species(self.species_guess)

        if self.micro_plot_id not in range(1, 6):
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'microplot ID', '1-5', self.micro_plot_id))

        if self.quarter not in range(1, 5):
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'quarter', '1-4', self.quarter))

        if self.scale not in [300, 1000]:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'scale', '300 or 1000', self.scale))

        if None != self.species_known and None != self.species_guess:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', ''))

        if None == self.species_known and None == self.species_guess:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', ''))

        if None != self.species_known:
            validation_errors += self.validate_species(self.species_known)

        if None != self.species_guess:
            validation_errors += self.validate_species(self.species_guess)

        species = (self.species_guess or self.species_known).lower()

        if (1 > self.diameter_breast_height) or (5 <= self.diameter_breast_height):
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'dbh', '>= 1 or < 5', self.diameter_breast_height))

        if not (self.diameter_breast_height * 10).is_integer():
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'dbh', '0.1 increments', self.diameter_breast_height))

        return validation_errors