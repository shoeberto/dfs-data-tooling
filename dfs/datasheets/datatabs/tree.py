from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError

class TreeTableTab(Tab):
    def __init__(self):
        super().__init__()

        self.tree_species = []


    def validate(self):
        validation_errors = []

        for species in self.tree_species:
            validation_errors += species.validate()

        return validation_errors


class TreeTableSpecies(Validatable):
    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.tree_number = None
        self.species_known = None
        self.species_guess = None
        self.diameter_breast_height = None
        self.live_or_dead = None
        self.comments = None


    def validate(self):
        validation_errors = []

        self.species_guess = self.override_species(self.species_guess)
        self.species_known = self.override_species(self.species_known)

        if self.micro_plot_id not in range(1, 6):
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'microplot ID', '1-5', self.micro_plot_id))

        if None != self.species_known and None != self.species_guess:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', ''))

        if None == self.species_known and None == self.species_guess:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', ''))

        if None != self.species_known:
            validation_errors += self.validate_species(self.species_known)

        if None != self.species_guess:
            validation_errors += self.validate_species(self.species_guess)

        species = (self.species_guess or self.species_known).lower()

        if 5 > self.diameter_breast_height:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'dbh', '> 5', self.diameter_breast_height))

        if not (self.diameter_breast_height * 10).is_integer():
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'dbh', '0.1 increments', self.diameter_breast_height))

        if self.live_or_dead not in ['L', 'D']:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'L or D', 'L or D', self.live_or_dead))

        return validation_errors