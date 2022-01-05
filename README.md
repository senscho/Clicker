Clicker: Menu / UI Element click emulation
========================

General
-------------------

  Clicker is a set of python scripts which enalbes clicking UI elements.
  It internally generates an applescript to perform a UI action.

  Examples of UI elements 
  
  * The [Ribbon Menu] found in Microsoft Applications
  * [MacOS Menu Bar] located at the top in MacOS

  With this program one can make

  * a custom API interface to virtually any Apps
  * Macros 

  [Ribbon Menu]: https://docs.microsoft.com/en-us/windows/win32/uxguide/cmd-ribbons
  [MacOS Menu Bar]: https://developer.apple.com/design/human-interface-guidelines/macos/menus/menu-bar-menus/

  Supported UI action in current version (can be extended easily, but time consuming)

  *	MS Word
  	* Menu Click
  		* File > Share > Send Document
  		* File > Share > Send PDF
  		* Insert > Break > Section Break (Next Page)
  		* Insert > Break > Section Break (Continuous)

  	* Ribbon Menu Click
  		* Home > Subscript
  		* Home > Superscript
  		* Home > Bullet
  		* Review > No mark up
  		* Review > All mark up
  		* Review > Previous Comment
  		* Review > Next Comment
  		* Review > Accept
  		* Review > Reject

Usage
-----

  **Menu Click**

  * MenuClicker.py "Program Name > Menu Name"

  Program Name / Menu Name are defined in UIOfficeMenu.py

  **UI Click**

  * UIClicker.py "Program Name > UI Name"

  Program Name / Menu Name are defined in UIOfficeRibbon.py


Stream Deck Integration
-----
  1. Make the python file executable
  	1. Change the first line of the script according to your system's python installation
  	1. make the file executable (i.e. use chmod or equivalent)
  1. In the stream deck, Use System -> Open feature.
  1. Select the python script and pass the argument. Example "/Users/jihyuncho/Library/Mobile Documents/com~apple~CloudDocs/_Personal/Coding/Clicker/MenuClicker.py" "Word > Send Document"
   	1. Use full path to your python script
  	1. Use double quotation for both the script and argument

  An example profile is located under Stream Deck Example directory.