from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Species
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError
from dfs.datasheets.datatabs.tabs import MissingSubplotValidationError
from dfs.datasheets.datatabs.tabs import FatalExceptionError
import numbers


class TreeTableTab(Tab):
    def __init__(self):
        super().__init__()

        self.tree_species = []


    def postprocess(self):
        self.check_primary_key_uniquness(self.tree_species, ['micro_plot_id', 'tree_number', 'get_species_known', 'get_species_guess'])


    def validate(self):
        validation_errors = []

        collected_subplots = []
        for species in self.tree_species:
            species.validate_attribute_type(['diameter_breast_height'], numbers.Number)

            try:
                validation_errors += species.validate()
            except Exception as e:
                validation_errors.append(FatalExceptionError(self.get_object_type(), e))

            collected_subplots.append(species.micro_plot_id)

        if set(collected_subplots) != set(range(1, 6)):
            validation_errors.append(MissingSubplotValidationError(self.get_object_type(), collected_subplots))

        return validation_errors


class TreeTableSpecies(Species):
    def __init__(self):
        super().__init__()

        self.micro_plot_id = None
        self.tree_number = None
        self.diameter_breast_height = None
        self.live_or_dead = None
        self.comments = None


    def validate(self):
        validation_errors = []

        self.species_guess = self.override_species(self.species_guess)
        self.species_known = self.override_species(self.species_known)

        if self.get_species_known() not in Tab.TREE_SPECIES:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species known', 'tree species', self.get_species_known()))

        if None != self.get_species_guess() and self.get_species_guess() not in Tab.TREE_SPECIES:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species guess', 'tree species', self.get_species_known()))

        if self.micro_plot_id not in range(1, Validatable.MAX_MICRO_PLOT_ID + 1):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'microplot ID', f'1-{Validatable.MAX_MICRO_PLOT_ID}', self.micro_plot_id))

        if None == self.species_known:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species known', 'non-empty', self.species_known))

        if None != self.species_known:
            validation_errors += self.validate_species(self.species_known)

        if None != self.species_guess:
            validation_errors += self.validate_species(self.species_guess)


        if None == self.diameter_breast_height:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'dbh', 'non-empty', self.diameter_breast_height))
        else: 
            if 5 > self.diameter_breast_height:
                validation_errors.append(FieldValidationError(self.get_object_type(), 'dbh', '> 5', self.diameter_breast_height))

            if not (self.diameter_breast_height * 10).is_integer():
                validation_errors.append(FieldValidationError(self.get_object_type(), 'dbh', '0.1 increments', self.diameter_breast_height))

        if self.live_or_dead not in ['L', 'D']:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'L or D', 'L or D', self.live_or_dead))

        return validation_errors