import sys
import glob
from os.path import basename
from dfs.datasheets import parsers
from dfs.datasheets import writer


def run(data_year, input_directory, output_directory):
    if 2013 == int(data_year):
        p = parsers.DatasheetParser2013()
    elif 2014 == int(data_year):
        p = parsers.DatasheetParser2014()
    elif 2015 == int(data_year):
        p = parsers.DatasheetParser2015()
    else:
        raise Exception(f'{data_year} is not a supported data year')

    w = writer.DatasheetWriter()

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


if 4 > len(sys.argv):
    raise Exception('Usage: python3 run.py [data year] [input file] [output directory]')


if '__main__' == __name__:
    run(sys.argv[1], sys.argv[2], sys.argv[3])
