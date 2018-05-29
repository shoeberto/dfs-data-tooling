from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationException
from dfs.datasheets.datatabs.tabs import FieldCountValidationException

class TreeTableTab(Tab):
    tree_species = []


    def validate(self):
        for species in self.tree_species:
            species.validate()


class TreeTableSpecies(Validatable):
    micro_plot_id = None
    tree_number = None
    species_known = None
    species_guess = None
    diameter_breast_height = None
    live_or_dead = None
    comments = None


    def validate(self):
        self.species_guess = self.override_species(self.species_guess)
        self.species_known = self.override_species(self.species_known)

        if self.micro_plot_id not in range(1, 6):
            raise FieldValidationException(self.__class__.__name__, 'microplot ID', '1-5', self.micro_plot_id)

        if None != self.species_known and None != self.species_guess:
            raise FieldValidationException(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', '')

        if None == self.species_known and None == self.species_guess:
            raise FieldValidationException(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', '')

        if None != self.species_known:
            self.validate_species(self.species_known)

        if None != self.species_guess:
            self.validate_species(self.species_guess)

        species = (self.species_guess or self.species_known).lower()

        if 5 > self.diameter_breast_height:
            raise FieldValidationException(self.__class__.__name__, 'dbh', '> 5', self.diameter_breast_height)

        if not (self.diameter_breast_height * 10).is_integer():
            raise FieldValidationException(self.__class__.__name__, 'dbh', '0.1 increments', self.diameter_breast_height)

        if self.live_or_dead not in ['L', 'D']:
            raise FieldValidationException(self.__class__.__name__, 'L or D', 'L or D', self.live_or_dead)
