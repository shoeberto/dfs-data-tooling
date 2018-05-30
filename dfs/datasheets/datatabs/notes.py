from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import FieldValidationException
from dfs.datasheets.datatabs.tabs import FieldCountValidationException

class NotesTab(Tab):
    seedlings = None
    seedlings_notes = None
    browsing = None
    browsing_notes = None
    indicators = None
    indicators_notes = None
    general_notes = None


    def validate(self):
        for field in ['seedlings', 'browsing', 'indicators']:
            self.validate_deer_impact(field)


    
    def validate_deer_impact(self, field):
        impact_rating = getattr(self, field)
        impact_notes = getattr(self, '{}_notes'.format(field))
        
        if None != impact_rating:
            if 1 > impact_rating or 5 < impact_rating:
                raise FieldValidationException(self.__class__.__name__, '{} impact rating'.format(field), '1-5', impact_rating)

            if None == impact_notes:
                raise FieldValidationException(self.__class__.__name__, '{} impact notes'.format(field), 'non-empy', impact_notes)

