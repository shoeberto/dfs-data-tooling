import sys
import glob
from dfs.datasheets import parser
from dfs.datasheets import writer


def run(data_year, input_directory, output_directory):
    p = parser.DatasheetParser2013()
    w = writer.DatasheetWriter()

    for filename in glob.glob('{}/*.xlsx'.format(input_directory)):
        datasheet = p.parse_datasheet(filename)
        w.write(datasheet, output_directory)


if 4 < len(sys.argv):
    raise Exception('Usage: python3 run.py [data year] [input file] [output directory]')


if '__main__' == __name__:
    run(sys.argv[1], sys.argv[2], sys.argv[3])