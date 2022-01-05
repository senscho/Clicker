#!/Users/jihyuncho/.pyenv/shims/python

from UIexplorer import codeBuilder
from UIOfficeMenu import *

import sys

# arg = "Word > Send PDF"
arg = sys.argv[1]

x = [x.strip() for x in arg.split('>')]

program = x[0]
cmd = office_ui[program][x[1]]

uiRef = []
uiRef.append('menu bar 1')
menuitems = [x.strip() for x in cmd.split('>')]

uiRef.append('menu bar item "{}"'.format(menuitems[0]))
uiRef.append('menu "{}"'.format(menuitems[0]))

for menu in menuitems[1:-1]:
	uiRef.append('menu item "{}"'.format(menu))
	uiRef.append('menu "{}"'.format(menu))

uiRef.append('menu item "{}"'.format(menuitems[-1]))

ui_string = ' of '.join(uiRef[-1::-1])

menuClicker = codeBuilder('window 1')
menuClicker.add_snippet("set frontmost to true")
menuClicker.add_snippet("click {}".format(ui_string))

# print(menuClicker.build_code())
menuClicker.run()