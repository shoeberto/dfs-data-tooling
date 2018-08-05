import csv
import numbers
from abc import ABC, abstractmethod


class Validatable(ABC):
    YES_NO_RESPONSE = ['Yes', 'No']
    YES_NO_TEXT = 'yes or no'
    SPECIES_OVERRIDES = {
        'acps': 'acpe',
        'acri': 'acru',
        'amar': 'amsp',
        'amel': 'amsp',
        'asca': 'acsa',
        'assp': 'aster',
        'beap': 'beal',
        'biv': 'bliv',
        'bolt': 'bliv',
        'caal': 'cato',
        'cape': 'caco',
        'euru': 'agal',
        'gass': 'grass',
        'gras': 'grass',
        'lsyp': 'lysp',
        'lysp2': 'lysp1',
        'muvi': 'mevi',
        'newi': 'thno',
        'osci': 'osvi',
        'posp': 'poten',
        'potr': 'potr5',
        'ptsp': 'ptaq',
        'sedg': 'sedge',
        'star': 'trbo',
        'sumac': 'rhus',
        'tica': 'tico',
        'toaf': 'taof',
        'trvo': 'trbo',
        'unk1': 'unka',
        'unk2': 'unkb',
        'vaac': 'vasp',
        'vapa': 'vasp',
        'rual': 'rusp',
        'vacc': 'vasp',
        'hupe': 'husp',
        'eram': 'ersp',
        'dead': 'dewo',
        'vivi': 'vitis',
        'lysp3': 'lysp',
        'befam': 'faga',
        'befa': 'faga',
        'prsp2': 'prsp'
    }

    COVER_SPECIES = None
    TREE_SPECIES = None
    SEEDLING_SPECIES = None

    MASTER_SPECIES_LIST = None


    def __init__(self):
        if None == Validatable.MASTER_SPECIES_LIST:
            with open('data/master_species_list.csv', 'r') as f:
                Validatable.MASTER_SPECIES_LIST = []
                Validatable.COVER_SPECIES = []
                Validatable.TREE_SPECIES = []
                Validatable.SEEDLING_SPECIES = []

                for row in csv.DictReader(f):
                    Validatable.MASTER_SPECIES_LIST.append(row['species_name'])

                    if 'TRUE' == row['cover']:
                        Validatable.COVER_SPECIES.append(row['species_name'])

                    if 'TRUE' == row['tree']:
                        Validatable.TREE_SPECIES.append(row['species_name'])

                    if 'TRUE' == row['seedling']:
                        Validatable.SEEDLING_SPECIES.append(row['species_name'])


    @abstractmethod
    def validate(self):
        pass


    def validate_mutually_required_fields(self, fields):
        field_none_values = {}

        for field in fields:
            field_none_values[field] = (None == getattr(self, field))

        if 1 < len(set(field_none_values.values())):
            return [RequiredFieldMismatchValidationError(self.get_object_type())]

        return []


    def override_scale(self, scale):
        if scale in [1000, 300]:
            return scale
        elif scale == 1:
            return 1000
        elif scale == 3:
            return 300

        raise Exception('{} is not a valid scale'.format(scale))


    def override_species(self, species):
        if None == species:
            return species

        try:
            species = species.lower().strip()
        except Exception:
            raise Exception("Cannot parse species code '{}'".format(species))

        if species in Validatable.SPECIES_OVERRIDES:
            return Validatable.SPECIES_OVERRIDES[species]

        return species

    
    def validate_species(self, species):
        if not isinstance(species, str):
            return [FieldValidationError(self.get_object_type(), 'species known or species guess', 'a text string', str(species))]

        species = species.lower().strip()

        if species not in Validatable.MASTER_SPECIES_LIST:
            return [SpeciesValidationError(self.get_object_type(), species)]

        if species[0:3] == 'unk':
            return [UnidentifiedUnknownError(self.get_object_type(), species)]

        return []


    def validate_attribute_type(self, attributes, data_type):
        errors = []
        for attribute in attributes:
            value = getattr(self, attribute)

            if None == value:
                continue

            if not isinstance(value, data_type):
                errors.append(InvalidDataTypeError(self.get_object_type(), attribute, str(data_type), str(type(value))))

        return errors


    def generate_primary_key(self, species, columns):
        values = []
        for column in columns:
            attr = getattr(species, column)

            if callable(attr):
                values.append(attr())
            else:
                values.append(attr)

        return ','.join([str(x) for x in values if x != None])


    def check_primary_key_uniquness(self, species, columns):
        try:
            primary_keys = [self.generate_primary_key(x, columns) for x in species]
        except Exception as e:
            print(FatalExceptionError(self.get_object_type(), e).get_message())
            raise Exception('Fatal error encountered generating primary key.')

        duplicates = set([x for x in primary_keys if primary_keys.count(x) > 1])

        if duplicates:
            for duplicate in duplicates:
                print(DuplicateRowValidationError(self.get_object_type(), duplicate).get_message())

            raise Exception('Cannot process file with duplicate cover species.')


    def get_object_type(self):
        return self.__class__.__name__


class Species(Validatable):
    def __init__(self):
        self.species_known = None
        self.species_guess = None


    def get_species_known(self):
        return self.override_species(self.species_known)


    def get_species_guess(self):
        return self.override_species(self.species_guess)


class Tab(Validatable):
    def __init__(self):
        pass
        

    @abstractmethod
    def postprocess(self):
        """Abstract method for performing post-processing."""


class ValidationError:
    def __init__(self):
        pass


    def get_message(self):
        return ''


class FieldValidationError(ValidationError):
    def __init__(self, object_type, field, expected, actual):
        self.object_type = object_type
        self.field = field
        self.expected = expected
        self.actual = actual


    def get_message(self):
        return "In {}: Expected value of field '{}' is '{}', got '{}'".format(self.object_type, self.field, self.expected, self.actual)


class FieldCountValidationError(ValidationError):
    def __init__(self, object_type, field, expected, actual):
        self.object_type = object_type
        self.field = field
        self.expected = expected
        self.actual = actual


    def get_message(self):
        return "In {}: Expected count of field '{}' to be {}, instead counted {}".format(self.object_type, self.field, self.expected, self.actual)


class DuplicateRowValidationError(ValidationError):
    def __init__(self, tab, combination):
        self.tab = tab
        self.combination = combination


    def get_message(self):
        return "In {}: Field values '{}' should be unique, but appears multiple times".format(self.tab, self.combination)


class RequiredFieldMismatchValidationError(ValidationError):
    def __init__(self, object_type):
        self.object_type = object_type


    def get_message(self):
        return 'In {}: mixture of empty and non-empty fields, where values are expected'.format(self.object_type)


class MissingSubplotValidationError(ValidationError):
    def __init__(self, object_type, collected_subplots):
        self.object_type =  object_type
        self.missing_subplots = [str(x) for x in (set(range(1, 6)) - set(collected_subplots))]

    
    def get_message(self):
        return 'In {}: data not collected for subplots {}'.format(self.object_type, ', '.join(self.missing_subplots))


class SpeciesValidationError(ValidationError):
    def __init__(self, object_type, species):
        self.object_type = object_type
        self.species = species


    def get_message(self):
        return "In {}: '{}' is not a valid species".format(self.object_type, self.species)


class InvalidDataTypeError(ValidationError):
    def __init__(self, object_type, field, expected_type, actual_type):
        self.object_type = object_type
        self.field = field
        self.expected_type = expected_type
        self.actual_type = actual_type


    def get_message(self):
        return "In {}: Field '{}' should be of type '{}', instead is '{}'".format(self.object_type, self.field, self.expected_type, self.actual_type)


class UnidentifiedUnknownError(ValidationError):
    def __init__(self, object_type, unknown_name):
        self.object_type = object_type
        self.unknown_name = unknown_name


    def get_message(self):
        return "In {}: Unidentified unknown species labeled as '{}'".format(self.object_type, self.unknown_name)


class FatalExceptionError(ValidationError):
    def __init__(self, object_type, exception):
        self.object_type = object_type
        self.exception = exception

    
    def get_message(self):
        return "Fatal exception thrown in {} with message: {}".format(self.object_type, str(self.exception))