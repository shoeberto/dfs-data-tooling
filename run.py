import sys
from dfs.datasheets import parser
from dfs.datasheets import writer


if 3 < len(sys.argv):
    raise Exception('Usage: python3 run.py [input file] [output directory]')


p = parser.DatasheetParser2013()
datasheet = p.parse_datasheet(sys.argv[1])

w = writer.DatasheetWriter()
w.write(datasheet, sys.argv[2])
