from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationError
from dfs.datasheets.datatabs.tabs import FieldCountValidationError

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


    def validate(self):
        validation_errors = []

        for field in ['seedlings', 'browsing', 'indicators']:
            validation_errors += self.validate_deer_impact(field)

        return validation_errors

    
    def validate_deer_impact(self, field):
        validation_errors = []
        impact_rating = getattr(self, field)
        impact_notes = getattr(self, '{}_notes'.format(field))
        
        if None != impact_rating:
            if 1 > impact_rating or 5 < impact_rating:
                validation_errors.append(FieldValidationError(self.__class__.__name__, '{} impact rating'.format(field), '1-5', impact_rating))

            if None == impact_notes:
                validation_errors.append(FieldValidationError(self.__class__.__name__, '{} impact notes'.format(field), 'non-empy', impact_notes))

        return validation_errors
