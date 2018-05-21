from dfs.datasheets.datatabs.tabs import Tab
from dfs.datasheets.datatabs.tabs import Validatable
from dfs.datasheets.datatabs.tabs import ValidationException


class WitnessTreeTab(Tab):
    witness_trees = []


    def validate(self):
        # TODO: validation code
        return


class WitnessTreeTabTree(Validatable):
    tree_number = None
    micro_plot_id = None
    species_known = None
    species_guess = None
    dbh = None
    live_or_dead = None
    azimuth = None
    distance = None


    def validate(self):
        # TODO: validation code
        return