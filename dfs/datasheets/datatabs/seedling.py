from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationException
from dfs.datasheets.datatabs.tabs import FieldCountValidationException

class SeedlingTable(Tab):
    seedling_species = []


    def validate(self):
        for species in self.seedling_species:
            species.validate()


class SeedlingSpecies(Validatable):
    micro_plot_id = None
    quarter = None
    scale = None
    species_known = None
    species_guess = None
    sprout = None
    zero_six_inches = None
    six_twelve_inches = None
    one_three_feet_total = None
    one_three_feet_browsed = None
    three_five_feet_total = None
    three_five_feet_browsed = None
    greater_five_feet_total = None
    greater_five_feet_browsed = None


    def validate(self):
        self.species_guess = self.override_species(self.species_guess)
        self.species_known = self.override_species(self.species_known)

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

        if self.sprout not in [0, 1]:
            raise FieldValidationException(self.__class__.__name__, 'sprout', '0 or 1', self.sprout)

        if None != self.zero_six_inches and 0 > self.zero_six_inches:
            raise FieldValidationException(self.__class__.__name__, '0-6"', '>= 0', self.zero_six_inches)

        if None != self.six_twelve_inches and 0 > self.six_twelve_inches:
            raise FieldValidationException(self.__class__.__name__, '6-12"', '>= 0', self.six_twelve_inches)

        if None != self.one_three_feet_total and 0 > self.one_three_feet_total:
            raise FieldValidationException(self.__class__.__name__, "1-3' Total", '>= 0', self.one_three_feet_total)

        if None != self.one_three_feet_browsed and 0 > self.one_three_feet_browsed:
            raise FieldValidationException(self.__class__.__name__, "1-3' Browsed", '>= 0', self.one_three_feet_browsed)

        if None != self.one_three_feet_browsed and None != self.one_three_feet_total:
            if self.one_three_feet_browsed > self.one_three_feet_total:
                raise FieldValidationException(self.__class__.__name__, "1-3' Browsed'", "<= 1-3' Total", self.one_three_feet_browsed)

        if None != self.three_five_feet_total and 0 > self.three_five_feet_total:
            raise FieldValidationException(self.__class__.__name__, "3-5' Total", '>= 0', self.three_five_feet_total)

        if None != self.three_five_feet_browsed and 0 > self.three_five_feet_browsed:
            raise FieldValidationException(self.__class__.__name__, "3-5' Browsed", '>= 0', self.three_five_feet_browsed)

        if None != self.three_five_feet_browsed and None != self.three_five_feet_total:
            if self.three_five_feet_browsed > self.three_five_feet_total:
                raise FieldValidationException(self.__class__.__name__, "3-5' Browsed", "<= 3-5' Total", self.one_three_feet_browsed)

        if None != self.greater_five_feet_total and 0 > self.greater_five_feet_total:
            raise FieldValidationException(self.__class__.__name__, ">5' Total", '>= 0', self.greater_five_feet_total)

        if None != self.greater_five_feet_browsed and 0 > self.greater_five_feet_browsed:
            raise FieldValidationException(self.__class__.__name__, ">5' Browsed", '>= 0', self.greater_five_feet_browsed)

        if None != self.greater_five_feet_browsed and None != self.greater_five_feet_total:
            if self.greater_five_feet_browsed > self.greater_five_feet_total:
                raise FieldValidationException(self.__class__.__name__, ">5' Browsed", "<= >5' Total", self.greater_five_feet_browsed)
