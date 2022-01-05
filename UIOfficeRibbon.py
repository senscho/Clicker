office_ui = {}


## MS Word
if True:
	word_ui = {}

	wRibbon = """
		if not exists scroll area 1 of tab group 1 of window 1 then click radio button "{0}" of tab group 1 of window 1
		delay 0.1
		if description of scroll area 1 of tab group 1 of window 1 does not contain "{0}" then click radio button "{0}" of tab group 1 of window 1
	"""

	word_ui['subscript'] = wRibbon.format('Home') + """
		checkbox "Subscript" of group 2 of scroll area 1 of tab group 1 of window 1
	"""

	word_ui['superscript'] = wRibbon.format('Home') + """
		checkbox "Superscript" of group 2 of scroll area 1 of tab group 1 of window 1
	"""

	word_ui['bullet'] = wRibbon.format('Home') + """
		click menu button "Bullets" of group 3 of scroll area 1 of tab group 1 of window 1
	"""

	word_ui['No Markup'] = wRibbon.format('Review') + """
	end tell

	ignoring application responses
		tell application "System Events"
			click pop up button 1 of group 6 of scroll area 1 of tab group 1 of window 1 of process "Word"
		end tell
	end ignoring

	delay 0.1
	do shell script "killall " & quote & "System Events" & quote
	delay 0.1

	tell application "System Events"
		click menu item "No Markup" of menu 1 of pop up button 1 of group 6 of scroll area 1 of tab group 1 of window 1 of process "Word"
	"""

	word_ui['All Markup'] = wRibbon.format('Review') + """
	end tell

	ignoring application responses
		tell application "System Events"
			click pop up button 1 of group 6 of scroll area 1 of tab group 1 of window 1 of process "Word"
		end tell
	end ignoring

	delay 0.1
	do shell script "killall " & quote & "System Events" & quote
	delay 0.1

	tell application "System Events"
		click menu item "All Markup" of menu 1 of pop up button 1 of group 6 of scroll area 1 of tab group 1 of window 1 of process "Word"
	"""

	word_ui['Prev Comment'] = wRibbon.format('Review') + """
		click button "Previous" of group 5 of scroll area 1 of tab group 1 of window 1
	"""

	word_ui['Next Comment'] = wRibbon.format('Review') + """
		click button "Next" of group 5 of scroll area 1 of tab group 1 of window 1
	"""

	word_ui['Accept'] = wRibbon.format('Review') + """
		click menu button "Accept" of group 8 of scroll area 1 of tab group 1 of window 1
	"""


	word_ui['Reject'] = wRibbon.format('Review') + """
		click menu button "Reject" of group 8 of scroll area 1 of tab group 1 of window 1
	"""

	office_ui['Word'] = word_ui