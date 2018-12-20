import sys
import argparse
import glob
from os.path import basename
from dfs.datasheets import parsers
from dfs.datasheets import writers
from dfs.datasheets.datatabs.tabs import Validatable


def run(data_year, input_directory, output_directory, data_type):
    if 'treatment' == data_type:
        Validatable.MAX_MICRO_PLOT_ID = 15

    p = get_parser(data_type, data_year)

    w = get_writer(data_type)

    for filename in sorted(glob.glob('{}/*.xlsx'.format(input_directory))):
        print("{}:".format(basename(filename)))

        try:
            datasheet = p.parse_datasheet(filename)

            datasheet.postprocess()
            datasheet.validate()

            try:
                w.write(datasheet, output_directory)
            except Exception:
                raise Exception('Encountered a fatal error writing the converted file. Validation errors may need to be fixed first.')

        except Exception as e:
            print('Fatal error encountered: {}'.format(str(e)))

        print('\n')


def get_parser(data_type, data_year):
    p = None

    if 'plot' == data_type:
        if 2013 == int(data_year):
            p = parsers.plots.DatasheetParser2013()
        elif 2014 == int(data_year):
            p = parsers.plots.DatasheetParser2014()
        elif 2015 == int(data_year):
            p = parsers.plots.DatasheetParser2015()
        elif 2016 == int(data_year):
            p = parsers.plots.DatasheetParser2016()
        elif 2017 == int(data_year):
            p = parsers.plots.DatasheetParser2017()
        else:
            raise Exception(f'{data_year} is not a supported data year')
    elif 'treatment' == data_type:
        if 2014 == data_year:
            p = parsers.treatments.TreatmentDatasheetParser2014()
        else:
            raise Exception(f'{data_year} is not a supported data year')
    elif 'superplot' == data_type:
        raise Exception('unsupported')

    return p


def get_writer(data_type):
    w = None

    if 'plot' == data_type:
        w = writers.plot.PlotDatasheetWriter()
    elif 'treatment' == data_type:
        w = writers.treatment.TreatmentDatasheetWriter()

    return w


parser = argparse.ArgumentParser(description='Validate and convert a set of data.')
parser.add_argument('-i', '--input-directory', help='Input directory.', required=True)
parser.add_argument('-o', '--output-directory', help='Output directory.', required=True)
parser.add_argument('-y', '--year', type=int, help='Data year.', required=True)
parser.add_argument('-t', '--data-type', help='Data type.', choices=['plot', 'treatment', 'superplot'], default='plot')

args = parser.parse_args()

if '__main__' == __name__:
    run(args.year, args.input_directory, args.output_directory, args.data_type)
