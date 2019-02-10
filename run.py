import sys
import os.path
import argparse
import glob
from os.path import basename
from dfs.datasheets import parsers
from dfs.datasheets import writers
from dfs.datasheets.datatabs.tabs import Validatable


def run(data_year, input_directory, output_directory, data_type):
    if 'treatment' == data_type:
        Validatable.MAX_MICRO_PLOT_ID = 15
    elif 'superplot' == data_type:
        Validatable.MAX_MICRO_PLOT_ID = 11

    p = get_parser(data_type, data_year)

    w = get_writer(data_type)

    for filename in sorted(glob.glob(f'{input_directory}/*.xlsx')):
        print(f'{basename(filename)}:')

        try:
            datasheet = p.parse_datasheet(filename)

            datasheet.postprocess()
            datasheet.validate()

            try:
                w.write(datasheet, output_directory)
            except Exception:
                raise Exception('Encountered a fatal error writing the converted file. Validation errors may need to be fixed first.')

        except Exception as e:
            print(f'Fatal error encountered: {str(e)}')

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
        elif 2015 == data_year:
            p = parsers.treatments.TreatmentDatasheetParser2015()
        elif 2016 == data_year:
            p = parsers.treatments.TreatmentDatasheetParser2016()
        elif 2017 == data_year:
            p = parsers.treatments.TreatmentDatasheetParser2017()
        else:
            raise Exception(f'{data_year} is not a supported data year')
    elif 'superplot' == data_type:
        if 2014 == data_year:
            p = parsers.superplots.SuperplotDatasheetParser2014()
        elif 2015 == data_year:
            p = parsers.superplots.SuperplotDatasheetParser2015()
        elif 2016 == data_year:
            p = parsers.superplots.SuperplotDatasheetParser2016()
        else:
            raise Exception(f'{data_year} is not a supported data year')

    return p


def get_writer(data_type):
    w = None

    if 'plot' == data_type:
        w = writers.plot.PlotDatasheetWriter()
    elif 'treatment' == data_type:
        w = writers.treatment.TreatmentDatasheetWriter()
    elif 'superplot' == data_type:
        w = writers.superplot.SuperplotDatasheetWriter()

    return w


parser = argparse.ArgumentParser(description='Validate and convert a set of data.')
parser.add_argument('-i', '--input-directory', help='Input directory.', required=True)
parser.add_argument('-o', '--output-directory', help='Output directory.', required=True)
parser.add_argument('-y', '--year', type=int, help='Data year.', required=True)
parser.add_argument('-t', '--data-type', help='Data type.', choices=['plot', 'treatment', 'superplot'], default='plot')

args = parser.parse_args()

if not os.path.isdir(args.input_directory):
    raise Exception(f"'{args.input_directory}' is not a directory")

if not os.path.exists(args.output_directory):
    os.makedirs(args.output_directory)
elif not os.path.isdir(args.output_directory):
    raise Exception(f"'{args.output_directory}' is not a valid output directory")

if '__main__' == __name__:
    run(args.year, args.input_directory, args.output_directory, args.data_type)
