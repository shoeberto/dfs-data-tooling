from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationException
from dfs.datasheets.datatabs.tabs import FieldCountValidationException

class CoverTableTab(Tab):
    cover_species = []


    def validate(self):
        for species in self.cover_species:
            species.validate()



class CoverSpecies(Validatable):
    DEER_INDICATOR_SPECIES = ['mara', 'maca', 'pobi', 'popu', 'trill', 'trer', 'trun', 'mevi']
    HEIGHT_OPTIONAL_SPECIES = ['dewo', 'moss']

    micro_plot_id = None
    quarter = None
    scale = None
    species_known = None
    species_guess = None
    percent_cover = None
    average_height = None
    count = None
    flower = None
    number_of_stems = None


    def validate(self):
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

        if self.percent_cover not in range(0, 10):
            raise FieldValidationException(self.__class__.__name__, 'percent cover', '0-9', self.percent_cover)

        if species not in self.HEIGHT_OPTIONAL_SPECIES and self.average_height not in range(1, 6):
            raise FieldValidationException(self.__class__.__name__, 'average height', '1-5', self.average_height)

        if species in self.DEER_INDICATOR_SPECIES:
            if None == self.count:
                raise FieldValidationException(self.__class__.__name__, 'count of species {}'.format(self.species_guess or self.species_known), 'non-empty', self.count)

            if None == self.flower:
                raise FieldValidationException(self.__class__.__name__, 'count of flowering for species {}'.format(self.species_guess or self.species_known), 'non-empty', self.flower)


        # TODO: flowering only required for trillium, but what species?



