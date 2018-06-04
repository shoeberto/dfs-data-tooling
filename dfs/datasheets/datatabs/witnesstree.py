from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError


class WitnessTreeTab(Tab):
    def __init__(self):
        super().__init__()

        self.witness_trees = []


    def validate(self):
        validation_errors = []

        if 3 > len(self.witness_trees):
            validation_errors.append(FieldCountValidationError(self.__class__.__name__, 'witness trees', 'at least 3', len(self.witness_trees)))

        for i in range(0, 3):
            if i + 1 != int(self.witness_trees[i].tree_number):
                validation_errors.append(FieldValidationError(self.__class__.__name__, 'witness tree number order', '1-3 contiguous numbering', self.witness_trees[i].tree_number))

        if 3 < len(self.witness_trees):
            for i in range(3, len(self.witness_trees)):
                if 0 == i % 2 and 2 != int(self.witness_trees[i]):
                    validation_errors.append(FieldValidationError(self.__class__.__name__, 'witness tree number order', '1-2 contiguous numering', self.witness_trees[i].tree_number))
                elif 1 == i % 2 and 1 != int(self.witness_trees[i]):
                    validation_errors.append(FieldValidationError(self.__class__.__name__, 'witness tree number order', '1-2 contiguous numering', self.witness_trees[i].tree_number))

        for witness_tree in self.witness_trees:
            validation_errors += witness_tree.validate()

        return validation_errors


class WitnessTreeTabTree(Validatable):
    def __init__(self):
        super().__init__()

        self.tree_number = None
        self.micro_plot_id = None
        self.species_known = None
        self.species_guess = None
        self.dbh = None
        self.live_or_dead = None
        self.azimuth = None
        self.distance = None


    def validate(self):
        validation_errors = []

        self.species_known = self.override_species(self.species_known)
        self.species_guess = self.override_species(self.species_guess)

        if int(self.tree_number) not in range(1, 4):
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'tree number', '1-3', self.tree_number))

        if int(self.micro_plot_id) not in range(1, 6):
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'micro plot ID', '1-5', self.micro_plot_id))

        if None != self.species_known and None != self.species_guess:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', ''))

        if None == self.species_known and None == self.species_guess:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'species known/species guess', 'one empty, one non-empty', ''))

        if None != self.species_known:
            validation_errors += self.validate_species(self.species_known)

        if None != self.species_guess:
            validation_errors += self.validate_species(self.species_guess)

        if self.dbh < 5:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'DBH', '> 5', self.dbh))

        if self.live_or_dead not in ['L', 'D']:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'live or dead', 'L or D', self.live_or_dead))

        if self.azimuth < 0 or self.azimuth > 359:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'azimuth', '0-359', self.azimuth))

        if None == self.distance or self.distance < 0:
            validation_errors.append(FieldValidationError(self.__class__.__name__, 'distance', '> 0', self.distance))

        return validation_errors