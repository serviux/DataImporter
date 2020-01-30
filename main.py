
from excel_helper import ExcelInterface
import os, os.path
import argparse
import sys

_desc = """
This program is used to interpret New England Tech's Smash League
Excel spreadsheet and store it into a json file. The json file is 
default created in the same directory as the excel file,
unless specified otherwise.
"""

ap = argparse.ArgumentParser(description=_desc)
ap.add_argument("input",  help ="Path to input spreadsheet")
ap.add_argument("--output", required=False, help = "Path to directory the file should be stored in")

args = ap.parse_args()

try:
    if args.input is None:
        raise IOError("path to excel sheet must be provided")
    elif not os.path.exists(args.input):
        raise IOError("File not found")

    if args.output is not None:
        if not os.path.exists(args.output):
            raise IOError("Path to output directory does not exist")
except IOError as e:
    print(e)
    sys.exit(0)

has_out_path = args.output is not None
xl = ExcelInterface(args.input)
xl.load_sheets()
xl.get_entries()
xl.find_scores()
xl.playersToJSON()
if has_out_path:
    xl.store_to_json(args.output)
else:
    head, tail = os.path.split(args.input)
    xl.store_to_json(head)






