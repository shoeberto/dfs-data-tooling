class GeneralTab:
    study_area = None
    plot_number = None
    deer_impact = None
    collection_date = None
    subplots = []
    auxillary_post_locations = []
    non_forested_azimuths = []


class PlotGeneralTab(GeneralTab):
    fenced_subplot_condition = None


class GeneralTabSubplot:
    latitude = None
    longitude = None
    micro_plot_id = None    # 1-5
    collected = None        # Yes/No
    fenced = None           # Yes/No
    azimuth = None          # 0-365
    distance = None         # Feet
    altitude = None         # Meters
    notes = None


class PlotGeneralTabSubplot(GeneralTabSubplot):
    re_monumented = None    # Yes/No
    forested = None         # Yes/No
    disturbance = None      # 0-1
    disturbance_type = None # 0-4; if disturbance = 0, must be 0; if disturbance = 1, must be 1-4
    lime = None             # Yes/No
    herbicide = None        # Yes/No



class TreatmentPlotGeneralPlotSubplot(GeneralTabSubplot):
    def __init__(self):
        # passthrough
        GeneralTabSubplot.__init__(self)


class FencedSubplotConditions:
    active_exclosure = None # text?
    repairs = None          # text?
    level = None            # 0-3
    notes = None            # text


class AuxillaryPostLocation:
    micro_plot_id = None
    post = None             # 1 or 2
    stake_type = None       # "Wooden" or "Rebar"
    azimuth = None          # 0 - 359
    distance = None         # Feet


class NonForestedAzimuths:
    micro_plot_id = None
    azimuth_1 = None
    azimuth_2 = None
    azimuth_3 = None
