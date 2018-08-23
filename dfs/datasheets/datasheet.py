TAB_NAME_GENERAL = 'General'
TAB_NAME_WITNESS_TREES = 'Witness_Trees'
TAB_NAME_NOTES = 'Notes'
TAB_NAME_TREE_TABLE = 'Tree_Table'
TAB_NAME_COVER_TABLE = 'Cover_Table'
TAB_NAME_SAPLING = 'Sapling_(1-5)'
TAB_NAME_SEEDLING = 'Seedling_(0-1)'


class Datasheet:
    def __init__(self):
        self.tabs = {
            TAB_NAME_GENERAL: None,
            TAB_NAME_WITNESS_TREES: None,
            TAB_NAME_NOTES: None,
            TAB_NAME_TREE_TABLE: None,
            TAB_NAME_COVER_TABLE: None,
            TAB_NAME_SAPLING: None,
            TAB_NAME_SEEDLING: None
        }


    def postprocess(self):
        for tab in self.tabs.values():
            tab.postprocess()


    def validate(self):
        file_validation_errors = []
        for tab in self.tabs.values():
            file_validation_errors += tab.validate()

        print('\n'.join([x.get_message() for x in file_validation_errors]))
