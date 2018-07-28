# from collections import defaultdict
# from pprint import pprint

from hachoir_metadata import metadata
from hachoir_core.cmd_line import unicodeFilename
from hachoir_parser import createParser

filename = './hsu-contributions.xls' 
filename, realname = unicodeFilename(filename), filename
parser = createParser(filename)
metalist = metadata.extractMetadata(parser).exportPlaintext()

print metalist[1:4]
