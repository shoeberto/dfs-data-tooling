import csv
from abc import ABC, abstractmethod

MASTER_SPECIES_LIST = None


class Validatable(ABC):
    YES_NO_RESPONSE = ['Yes', 'No']
    YES_NO_TEXT = 'yes or no'


    @abstractmethod
    def validate(self):
        pass


    def validate_mutually_required_fields(self, fields):
        field_none_values = {}

        for field in fields:
            field_none_values[field] = (None == getattr(self, field))

        if 1 < len(set(field_none_values.values())):
            return [RequiredFieldMismatchValidationError()]

        return []


    def override_species(self, species):
        if None == species:
            return species

        species = species.lower().strip()

        if 'vapa' == species:
            return 'vasp'
        
        if 'rual' == species:
            return 'rusp'

        if 'vacc' == species:
            return 'vasp'

        if 'hupe' == species:
            return 'husp'

        return species

    
    def validate_species(self, species):
        global MASTER_SPECIES_LIST

        if None == MASTER_SPECIES_LIST:
            with open('data/master_species_list.csv', 'r') as f:
                MASTER_SPECIES_LIST = [row['species_name'] for row in csv.DictReader(f)]

        if species.lower().strip() not in MASTER_SPECIES_LIST:
            return [SpeciesValidationError(self.__class__.__name__, species)]

        return []


class Tab(Validatable):
    def __init__(self):
        pass


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


class RequiredFieldMismatchValidationError(ValidationError):
    def __init__(self, object_type):
        self.object_type = object_type


    def get_message(self):
        return 'In {}: mixture of empty and non-empty fields, where values are expected'.format(self.object_type)


class SpeciesValidationError(ValidationError):
    def __init__(self, object_type, species):
        self.object_type = object_type
        self.species = species


    def get_message(self):
        return "In {}: '{}' is not a valid species".format(self.object_type, self.species)