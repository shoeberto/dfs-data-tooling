from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Species
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError
from dfs.datasheets.datatabs.tabs import FatalExceptionError
import numbers


class WitnessTreeTab(Tab):
    def __init__(self):
        super().__init__()

        self.witness_trees = []


    def postprocess(self):
        pass


    def validate(self):
        validation_errors = []

        if 3 > len(self.witness_trees):
            validation_errors.append(FieldCountValidationError(self.get_object_type(), 'witness trees', 'at least 3', len(self.witness_trees)))

        for i in range(0, 3):
            if i >= len(self.witness_trees):
                continue

            if i + 1 != int(self.witness_trees[i].tree_number):
                validation_errors.append(FieldValidationError(self.get_object_type(), 'witness tree number order', '1-3 contiguous numbering', self.witness_trees[i].tree_number))

        if 3 < len(self.witness_trees):
            for i in range(3, len(self.witness_trees)):
                if 0 == i % 2 and 2 != int(self.witness_trees[i].tree_number):
                    validation_errors.append(FieldValidationError(self.get_object_type(), 'witness tree number order', '1-2 contiguous numering', self.witness_trees[i].tree_number))
                elif 1 == i % 2 and 1 != int(self.witness_trees[i].tree_number):
                    validation_errors.append(FieldValidationError(self.get_object_type(), 'witness tree number order', '1-2 contiguous numering', self.witness_trees[i].tree_number))

        for witness_tree in self.witness_trees:
            validation_errors += witness_tree.validate_attribute_type(['dbh', 'azimuth', 'distance'], numbers.Number)
            validation_errors += witness_tree.validate_attribute_type(['live_or_dead'], str)

            try:
                validation_errors += witness_tree.validate()
            except Exception as e:
                validation_errors.append(FatalExceptionError(self.get_object_type(), e))

        return validation_errors


class WitnessTreeTabTree(Species):
    def __init__(self):
        super().__init__()

        self.tree_number = None
        self.micro_plot_id = None
        self.dbh = None
        self.live_or_dead = None
        self.azimuth = None
        self.distance = None


    def validate(self):
        validation_errors = []

        self.species_known = self.override_species(self.species_known)
        self.species_guess = self.override_species(self.species_guess)

        if int(self.tree_number) not in range(1, 4):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'tree number', '1-3', self.tree_number))

        if self.micro_plot_id not in range(1, Validatable.MAX_MICRO_PLOT_ID + 1):
            validation_errors.append(FieldValidationError(self.get_object_type(), 'micro plot ID', f'1-{Validatable.MAX_MICRO_PLOT_ID}', self.micro_plot_id))

        if None == self.species_known:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'species known', 'non-empty', self.species_known))

        if None != self.species_known:
            validation_errors += self.validate_species(self.species_known)

        if None != self.species_guess:
            validation_errors += self.validate_species(self.species_guess)

        if None == self.dbh or self.dbh < 5:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'DBH', '> 5', self.dbh))

        if None == self.live_or_dead or self.live_or_dead not in ['L', 'D']:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'live or dead', 'L or D', self.live_or_dead))

        if None == self.azimuth or self.azimuth < 0 or self.azimuth > 359:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'azimuth', '0-359', self.azimuth))

        if None == self.distance or self.distance < 0:
            validation_errors.append(FieldValidationError(self.get_object_type(), 'distance', '> 0', self.distance))

        return validation_errors