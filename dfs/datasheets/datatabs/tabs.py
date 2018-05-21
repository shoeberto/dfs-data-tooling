from abc import ABC, abstractmethod


class Validatable(ABC):
    @abstractmethod
    def validate(self):
        pass



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