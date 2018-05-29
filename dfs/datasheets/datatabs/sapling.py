from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationException
from dfs.datasheets.datatabs.tabs import FieldCountValidationException

class SaplingTab(Tab):
    sapling_species = []


    def validate(self):
        for species in self.sapling_species:
            species.validate()


class SaplingSpecies(Validatable):
    micro_plot_id = None
    sapling_number = None
    quarter = None
    scale = None
    species_known = None
    species_guess = None
    diameter_breast_height = None


    def validate(self):
        self.species_known = self.override_species(self.species_known)
        self.species_guess = self.override_species(self.species_guess)

        if self.micro_plot_id not in range(1, 6):
            raise FieldValidationException(self.__class__.__name__, 'microplot ID', '1-5', self.micro_plot_id)

        if self.quarter not in range(1, 5):
            raise FieldValidationException(self.__class__.__name__, 'quarter', '1-4', self.quarter)

        if self.scale not in [300, 1000]:
            raise FieldValidationException(self.__class__.__name__, 'scale', '300 or 1000', self.scale)

        if None != self.species_known and None != self.species_guess:
            raise FieldValidationException(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', '')

        if None == self.species_known and None == self.species_guess:
            raise FieldValidationException(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', '')

        if None != self.species_known:
            self.validate_species(self.species_known)

        if None != self.species_guess:
            self.validate_species(self.species_guess)

        species = (self.species_guess or self.species_known).lower()

        if (1 > self.diameter_breast_height) or (5 <= self.diameter_breast_height):
            raise FieldValidationException(self.__class__.__name__, 'dbh', '>= 1 or < 5', self.diameter_breast_height)

        if not (self.diameter_breast_height * 10).is_integer():
            raise FieldValidationException(self.__class__.__name__, 'dbh', '0.1 increments', self.diameter_breast_height)