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
            raise RequiredFieldMismatchValidationException(self.__class__.__name__)

    
    def validate_species(self, species):
        global MASTER_SPECIES_LIST

        if None == MASTER_SPECIES_LIST:
            with open('data/master_species_list.csv', 'r') as f:
                MASTER_SPECIES_LIST = [row['species_name'] for row in csv.DictReader(f)]

        # TODO: where are these?
        if species.lower().strip() in ['vacc', 'hupe']:
            return

        if species.lower().strip() not in MASTER_SPECIES_LIST:
            raise SpeciesValidationException(self.__class__.__name__, species)


class Tab(Validatable):
    def __init__(self):
        return


class ValidationException(Exception):
    def __init__(self, message):
        super().__init__(message)


class FieldValidationException(ValidationException):
    def __init__(self, object_type, field, expected, actual):
        super().__init__("In {}: Expected value of field '{}' is '{}', got '{}'".format(object_type, field, expected, actual))


class FieldCountValidationException(ValidationException):
    def __init__(self, object_type, field, expected, actual):
        super().__init__("In {}: Expected count of field '{}' to be {}, instead counted {}".format(object_type, field, expected, actual))


class RequiredFieldMismatchValidationException(ValidationException):
    def __init__(self, object_type):
        super().__init__('In {}: mixture of empty and non-empty fields, where values are expected'.format(object_type))


class SpeciesValidationException(ValidationException):
    def __init__(self, object_type, species):
        super().__init__("In {}: '{}' is not a valid species".format(object_type, species))