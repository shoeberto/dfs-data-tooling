from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Species
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError
from dfs.datasheets.datatabs.tabs import MissingSubplotValidationError
from dfs.datasheets.datatabs.tabs import DuplicateRowValidationError
from dfs.datasheets.datatabs.tabs import FatalExceptionError
import numbers


class CoverTableTab(Tab):
    def __init__(self):
        super().__init__()

        self.cover_species = []


    def postprocess(self):
        self.check_primary_key_uniquness(self.cover_species, ['micro_plot_id', 'quarter', 'scale', 'get_species_known', 'get_species_guess'])

        species_equal = (lambda lhs, rhs: lhs.micro_plot_id == rhs.micro_plot_id and lhs.quarter == rhs.quarter and lhs.get_species_known() == rhs.get_species_known())

        three_hundredth_species_list = [x for x in self.cover_species if x.scale == 300]

        new_species = []
        for milacre_species in [x for x in self.cover_species if x.scale == 1000]:
            milacre_species_in_three_hundredth = False
            for three_hundredth_species in three_hundredth_species_list:
                if species_equal(milacre_species, three_hundredth_species):
                    milacre_species_in_three_hundredth = True
                    three_hundredth_species.fill_empty_values_from_cover_species(milacre_species)

            if not milacre_species_in_three_hundredth and milacre_species not in Tab.TREE_SPECIES:
                print('Copying species {},{},{},{} from milacre to 300th acre.'.format(
                    milacre_species.micro_plot_id,
                    milacre_species.quarter,
                    milacre_species.scale,
                    milacre_species.get_species_known()
                ))

                species = CoverSpecies()
                species.micro_plot_id = milacre_species.micro_plot_id
                species.quarter = milacre_species.quarter
                species.scale = 300
                species.species_known = milacre_species.species_known
                species.species_guess = milacre_species.species_guess
                species.fill_empty_values_from_cover_species(milacre_species)
                new_species.append(species)

        self.cover_species = self.cover_species + new_species

        for species in self.cover_species:
            species.postprocess()


    def validate(self):
        validation_errors = []

        collected_subplots = []
        primary_keys = []
        for species in self.cover_species:
            species.validate_attribute_type([
                'percent_cover',
                'average_height',
                'count',
                'flower',
                'number_of_stems'
            ], numbers.Number)

            try:
                validation_errors += species.validate()
            except Exception as e:
                validation_errors.append(FatalExceptionError(self.get_object_type(), e))

            collected_subplots.append(species.micro_plot_id)

        if sorted(set(collected_subplots)) != sorted(set(range(1, 6))):
            validation_errors.append(MissingSubplotValidationError(self.get_object_type(), collected_subplots))

        return validation_errors


    def get_recorded_subplots(self):
        subplots = []

        for species in self.cover_species:
            subplots.append(species.micro_plot_id)

        return set(subplots)


class CoverSpecies(Species):
    DEER_INDICATOR_SPECIES = ['mara', 'maca', 'pobi', 'popu', 'trill', 'trer', 'trun', 'mevi']
    TRILLIUM_SPECIES = ['trer', 'trun', 'trill']
    HEIGHT_OPTIONAL_SPECIES = ['dewo', 'moss', 'bliv', 'bodt', 'rock', 'root', 'road', 'trash', 'water', 'trail']
    COVER_SPECIES_OVERRIDES = {
        'snag': 'bodt',
        'prsp': 'prena'
    }
    

    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.quarter = None
        self.scale = None
        self.percent_cover = None
        self.average_height = None
        self.count = None
        self.flower = None
        self.number_of_stems = None


    def fill_empty_values_from_cover_species(self, source_species):
        for field in ['percent_cover', 'average_height', 'flower', 'number_of_stems']:
            if None == getattr(self, field) and None != getattr(source_species, field):
                print(f'Copying field {field} from {source_species.micro_plot_id},{source_species.quarter},{source_species.scale},{source_species.get_species_known()} to {self.micro_plot_id},{self.quarter},{self.scale},{self.get_species_known()}') 
                setattr(self, field, getattr(source_species, field))

        if None == self.percent_cover:
            return

        if 0 < self.percent_cover or self.get_species_known() in CoverSpecies.DEER_INDICATOR_SPECIES:
            print(f'Copying count from {source_species.micro_plot_id},{source_species.quarter},{source_species.scale},{source_species.get_species_known()} to {self.micro_plot_id},{self.quarter},{self.scale},{self.get_species_known()}') 
            self.count = source_species.count


    def get_species_known(self):
        return self.override_species(self.species_known)


    def get_species_guess(self):
        return self.override_species(self.species_guess)


    def override_species(self, species):
        if None == species:
            return None

        try:
            if species.lower() in CoverSpecies.COVER_SPECIES_OVERRIDES:
                return CoverSpecies.COVER_SPECIES_OVERRIDES[species.lower()]
        except Exception:
            raise Exception("Cannot parse species code '{}'".format(species))

        return super().override_species(species)


    def postprocess(self):
        if 'gapr' == self.get_species_known() and None == self.average_height:
            self.average_height = 1


    def validate(self):
        validation_errors = []

        self.species_known = self.override_species(self.species_known)
        self.species_guess = self.override_species(self.species_guess)

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

        if self.get_species_known() not in Validatable.COVER_SPECIES:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species known', 'cover species', self.get_species_known()))

        if None != self.get_species_guess() and self.get_species_guess() not in Validatable.COVER_SPECIES:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species guess', 'cover species', self.get_species_guess()))

        if self.percent_cover not in range(0, 10):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'percent cover of species {}'.format(self.species_known), '0-9', self.percent_cover))

        if self.species_known not in (CoverSpecies.HEIGHT_OPTIONAL_SPECIES + Tab.TREE_SPECIES) and self.average_height not in range(1, 6):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'average height of species {}'.format(self.species_known), '1-5', self.average_height))

        if self.species_known in CoverSpecies.DEER_INDICATOR_SPECIES:
            if None == self.count:
                validation_errors.append(FieldValidationError(self.get_object_type(), 'count of species {}'.format(self.species_guess or self.species_known), 'non-empty', self.count))
            elif 0 == self.count:
                validation_errors.append(FieldValidationError(self.get_object_type(), 'count of species {}'.format(self.species_guess or self.species_known), 'non-zero', self.count))

            if None == self.flower:
                validation_errors.append(FieldValidationError(self.get_object_type(), 'count of flowering for species {}'.format(self.species_guess or self.species_known), 'non-empty', self.flower))

        if self.species_known not in CoverSpecies.TRILLIUM_SPECIES and None != self.number_of_stems:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'nstem', 'empty for non-trillium species', self.number_of_stems))

        if self.species_known in CoverSpecies.TRILLIUM_SPECIES and None == self.number_of_stems:
            self.number_of_stems = 0

        return validation_errors