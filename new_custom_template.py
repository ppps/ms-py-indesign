#!/usr/bin/env python3

import json
from pathlib import Path
import sys

masters_file = Path('masters.json')

master_dict = json.loads(masters_file.read_text(encoding='utf-8'))
master_names = '# ' + '\n# '.join(n for n in master_dict)

preamble = f'''\
# This is a template for a new custom edition.
#
# The first line should be a title, ideally starting with
# the date in YYYY-MM-DD format (ie 2018-07-18) and then
# a description (ie 2018-07-14 Durham Miners’ Gala)
#
# Then type a series of lines describing a page or spread in
# the edition, starting with a page number, one or more
# spaces, then the name of a master page template.
#
# For example, to create the front page, type:
# 1     News-Front
#
# And a features spread:
#
# 8     Feat-Base-S
#
# New masters in the InDesign file need to be added to the
# masters.json file before you can generate them. These are
# the master pages that have entries in the masters.json file:
#
{master_names}
'''

print(preamble)
