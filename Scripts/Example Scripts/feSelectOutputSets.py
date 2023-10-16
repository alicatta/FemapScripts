# Example code to demonstrate FEMAP API programmed through Python
# 

# Import Python modules 
import pythoncom
import PyFemap
from PyFemap import constants as feConstants #Sets femap constants to constants object so calling constants in constants.FCL_BLACK, not PyFemap.constants.FCL_BLACK
import sys

import win32com.client as win32
win32.gencache.is_readonly=False

# Connect to FEMAP 
try:
    existObj = pythoncom.connect(PyFemap.model.CLSID) # Grabs active model
    App = PyFemap.model(existObj)
except:
    sys.exit("femap not open") # Exits program if there is no active femap model

rc = App.feAppMessage(0,"Python API Started")

# Main funciton loop - invoked elsewhere
def feSelectOutputs(title = "Select FEMAP Outputs") :

	## Setup initial objects and UI structural components ##
	
	# Import additional Modules
	import tkinter as tk 
	
	# Setup FEMAP objects
	osetSet = App.feSet
	rbo = App.feResults
	
	# Setup root (master window)
	root = tk.Tk()
	root.title(title)
	root.geometry("350x440")
	
	# Frame padding factor for convenient formatting
	pad = 5
	
	# Setup Frames to house all other Widgets
	frameInfo = tk.Frame(root) # sets "root" as master of "frameInfo"
	# Specify where the frame is placed, how it fills the master,
	# expansion behavior as the window is resized, and any padding
	frameInfo.pack(side = tk.TOP, fill = tk.BOTH, expand = False, padx = pad, pady = pad)
	
	frameLabel = tk.LabelFrame(root, text = "Output Sets")
	frameLabel.pack(side = tk.TOP, fill = tk.BOTH, expand = True, padx = pad, pady = pad)
	
	frameTop = tk.Frame(frameLabel)#.pack(side = tk.TOP)
	frameTop.pack(side = tk.TOP, fill = tk.BOTH, expand = False, padx = pad, pady = pad)
	
	frameMid = tk.Frame(frameLabel)#.pack(side = tk.TOP)
	frameMid.pack(side = tk.TOP, fill = tk.BOTH, expand = True, padx = pad, pady = pad)
	
	frameBottom = tk.Frame(root)
	frameBottom.pack(side = tk.TOP, fill = tk.BOTH, expand = False, padx = pad, pady = pad)
	
	## Populate UI with Widgets ##
	
	# Display Model info in top portion of Dialogue
	strInfo = tk.StringVar()
	labelInfo = tk.Label(frameInfo, textvariable = strInfo)
	strInfo.set("Model -\t" + App.ModelName)
	labelInfo.pack(side = tk.LEFT, fill = tk.X, expand = False)
	
	# Display counter for number of Output Sets selected 
	strCount = tk.StringVar()
	labelCount = tk.Label(frameBottom, textvariable = strCount)
	strCount.set("None selected.")
	labelCount.pack(side = tk.LEFT, fill = tk.X, expand = False)
	
	# Define Listbox and Output Set IDs to populate it
	Lb1 = tk.Listbox(frameMid, selectmode = tk.MULTIPLE, yscrollcommand = True)
	osetSet.AddAll(feConstants.FT_OUT_CASE)
	
	# Bring Output Set IDs from FEMAP into Python 
	rc = osetSet.CopyToClipboard(True)
	if osetSet.Count() == 0 :
		# No Output Sets exist; exit
		return
	if rc == -1 :
		# Successfully populated and copied Output Set Set to clipboard;
		# Paste into Python as List
		osetIDs = list( map( int, root.selection_get(selection = "CLIPBOARD").splitlines() ) )
	else :
		# Output Sets exist, but did not copy; add IDs individually through Set
		osetIDs = []
		osetSet.Reset()
		while osetSet.Next() > 0 :
			osetIDs.append(osetSet.CurrentID)
		
	# Create list to house formatted titles of all Output Sets
	osetTitles = []
	# Dictionary to get Output Set ID from formatted title 
	osetDict = dict()
	# List to house results to return
	retVal = []
	
	# Create List of Output Set titles and IDs as "[ID]..[Title]"
	for ID in osetIDs :
		tstr = str(ID) + ".." + rbo.SetTitle(ID)[1]
		osetTitles.append( tstr )
		osetDict[tstr] = ID
	
	# Populate Listbox 
	for i in osetTitles: 
		Lb1.insert(tk.END, i)
	# Pack Listbox into corresponding Frame
	Lb1.pack(side = tk.LEFT, fill = tk.BOTH, expand = True)
	
	# Add Scroll Bar to Listbox Widget
	sb1 = tk.Scrollbar(frameMid, orient = tk.VERTICAL)
	sb1.pack(side = tk.LEFT, fill = tk.Y)
	
	Lb1.config(yscrollcommand = sb1.set)
	sb1.config(command = Lb1.yview)
	
	## Event handling ##
	
	# Event handling Functions 
	def cbButtonAll()  : # Select all Output Sets
		Lb1.select_set(0, tk.END)
		cbMouseClick(-1)
		
	def cbButtonNone() : # Clear any existing user selection
		Lb1.selection_clear(0, tk.END)
		cbMouseClick(-1)
		
	def cbButtonOK() : # Proceed with current user selection
		
		# Query Listbox for user selection
		for i in Lb1.curselection() : 
			# Append corresponding Output Set ID to return value
			retVal.append( osetDict[ Lb1.get(i) ] )
		
		# Exit GUI
		root.destroy()
		
	def cbButtonCancel() : # Exit dialog with no selection 
		
		root.destroy()
		return	
		
	def cbMouseClick(event) : # Respond to mouse clicks during runtime
		sel = Lb1.curselection()
		
		if len(sel) == 0 :
			strCount.set("None selected.")
		else :
			strCount.set( str(len(sel)) + " Output Sets selected.")
			
	def cbAddButton() : # Add a new button to UI
		# Purely an example of how the UI can build itself during
		# live runtime. Adds a button to the bottom Frame.
		newButton = tk.Button(frameBottom,text = "!", width = 2)
		newButton.pack(side = tk.RIGHT)
	
	## Create button Widgets ##
	
	# Button to select all Output Sets
	buttonAll = tk.Button(frameTop, text = "All", command = cbButtonAll, width = 5)
	buttonAll.pack(side = tk.LEFT, padx = 2)
	
	# Button to clear user selection 
	buttonNone = tk.Button(frameTop, text = "None", command = cbButtonNone, width = 5)
	buttonNone.pack(side = tk.LEFT, padx = 2)
	
	# Button to exit without making any selection
	buttonCancel = tk.Button(frameBottom, text = "Cancel", command = cbButtonCancel, width = 10)
	buttonCancel.pack(side = tk.RIGHT, padx = 2)
	
	# Button to proceed with current selection
	buttonOK = tk.Button(frameBottom, text = "Ok", command = cbButtonOK, width = 10)
	buttonOK.pack(side = tk.RIGHT, padx = 2)
	
	# Example of how UI can be built in real-time. No functional purpose
	#buttonAddButton = tk.Button(frameBottom, text = "Add", command = cbAddButton, width = 5)
	#buttonAddButton.pack(side = tk.RIGHT)
	
	# Connect mouse click event with custom mouse click function
	root.bind("<Button-1>", cbMouseClick)

	# Invoke main dialog 
	root.mainloop()
	
	# Exit Function and return value "retVal"
	return retVal
	
# Invoke main function
osetList = feSelectOutputs()

App.feAppMessage(0,"Python API Finished")