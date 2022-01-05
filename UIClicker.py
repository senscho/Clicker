#!/Users/jihyuncho/.pyenv/shims/python

from UIexplorer import codeBuilder
from UIOfficeRibbon import *

import sys

# arg = "Word > No Markup"
arg = sys.argv[1]

x = [x.strip() for x in arg.split('>')]

program = x[0]
cmd = office_ui[program][x[1]]

menuClicker = codeBuilder('window 1',process_name=program)
menuClicker.add_snippet("set frontmost to true")
menuClicker.add_snippet(cmd)
menuClicker.run()