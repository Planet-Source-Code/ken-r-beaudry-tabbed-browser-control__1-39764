Tabbed Browser Control - By Ken Beaudry
=======================================

This control was written because I started doing a couple of apps that were all using a tabbed browser interface, and I wanted to have a simple drop-in component that I could use to speed up development.

Notes:
======
If you Right-Click on a tab a popupmenu will be displayed.  Currently the only option is to delete the current tab.
There are updates coming in the future.
	- Owner drawn tab control (to replace MS Common controls)
	- Saving and loading current Tabs


Usage:

1) Draw the control on your target form
2) Set appropriate Properties (see below)
3) Initialize the control
		eg:
			Private Sub Form_Load()
			
			TBBrowser.InitControl "www.msn.com"
			
			End Sub

Methods:
========
NewTab(URL As String, Optional Options As Integer)
	- Creates a new browser tab and loads the specified URL
SelectTab (index As Integer)
	- Sets the specified tab to current and brings it's browser to the front
Back
	- Navigates the current browser back 1 page
Forward
	- Navigates the current browser forward 1 page	
Refresh
	- refreshed the current browser
Stop1
	- stops the current browser
Home
	- sends the current browser to the homepage specified in the homepage property
DeleteTab
	- deletes the currently selected tab






Properties:
===========
CustomError			- URL to your Custom Nav error page (if any)
StatusVisable			- Toggles wheather the status bar is visable
CurrentAddress			- Holds The URL if the displayed page
numtabs				- Holds the number of visable tabs
CurrentBrowser			- Holds the currently selected Browser Tab
Popups				- Sets whether or not to display popups
homepage			- Sets/Stores current homepage