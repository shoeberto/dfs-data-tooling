from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError
from dfs.datasheets.datatabs.tabs import FatalExceptionError
import numbers


class NotesTab(Tab):
    def __init__(self):
        super().__init__()

        self.seedlings = None
        self.seedlings_notes = None
        self.browsing = None
        self.browsing_notes = None
        self.indicators = None
        self.indicators_notes = None
        self.general_notes = None


    def postprocess(self):
        pass


    def validate(self):
        validation_errors = []

        validation_errors += self.validate_attribute_type(['seedlings', 'browsing', 'indicators'], numbers.Number)
        validation_errors += self.validate_attribute_type(['seedlings_notes', 'browsing_notes', 'indicators_notes'], str)

        try:
            for field in ['seedlings', 'browsing', 'indicators']:
                validation_errors += self.validate_deer_impact(field)
        except Exception as e:
            validation_errors.append(FatalExceptionError(self.get_object_type(), e))

        return validation_errors

    
    def validate_deer_impact(self, field):
        validation_errors = []
        impact_rating = getattr(self, field)
        impact_notes = getattr(self, '{}_notes'.format(field))
        
        if None != impact_rating:
            if 1 > impact_rating or 5 < impact_rating:
                validation_errors.append(FieldValidationError(self.get_object_type(), '{} impact rating'.format(field), '1-5', impact_rating))

            if None == impact_notes:
                validation_errors.append(FieldValidationError(self.get_object_type(), '{} impact notes'.format(field), 'non-empy', impact_notes))

        return validation_errors
