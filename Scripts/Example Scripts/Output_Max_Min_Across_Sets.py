# Example code to demonstrate FEMAP API programmed through Python
# This example outputs either the maximum or minimum value of chosen
# Output Vectors across chosen Output Sets 

# Import Python Modules to be used
import pythoncom
import PyFemap
from PyFemap import constants # To keep syntax similar to FEMAP API
import sys

# COM / OLE Objects to communicate across programs
import win32com.client as win32
win32.gencache.is_readonly=False # Enable writing

# Connect Python to active FEMAP session 
try:
    existObj = pythoncom.connect(PyFemap.model.CLSID) # Grabs active model
    App = PyFemap.model(existObj)
except: # Exit message if connection unsuccessful
    sys.exit("femap not open") #Exits program if there is no active femap model

# Print connection confirmation in FEMAP 
rc = App.feAppMessage(0,"Python API Started")

# Function to quickly import FEMAP Set as Python List 
def feSetIn(feSet_raw) :
	""" Import FEMAP Set as Python Set by copying and pasting from the Clipboard"""

	# Copy specified set to clipboard 
	rc = feSet_raw.CopyToClipboard(True)
	
	# Attempt to paste from clipboard, import from tkinter module if
	# paste attempt fails; attempts 3 times before failing function call
	attempts = 0
	while attempts < 3 :
		try :
			# TK().selection_get() calls paste from clipboard; 
			# list, map, int, splitlines() are all Python calls to parse 
			# resulting string
			return list( map( int, Tk().selection_get(selection = "CLIPBOARD").splitlines() ) )	
			
			# Function will exit after "return" call 
			
		except Exception :
		
			# Import TK to access clipboard
			from tkinter import Tk
			attempts += 1
	
	# if no "return" call, function failed: notify user, and return nothing
	App.feAppMessage(constants.FCM_NORMAL,"Failed to copy Set ID " + \
	str(feSet_raw.ID) + "to Clipboard")
	return []

# Main funciton loop - invoked elsewhere
def main() :

	# Prepare any FEMAP objects needed
	osetSet = App.feSet
	vecSet = App.feSet
	v = App.feOutput
	res = App.feResults
	table = App.feDataTable
	
	# List of Lists containing data to be written to Table
	dataOut = [[], [], [], []] # [[eID] [Vector] [osetID] [Value]]
	
	# Prompt User to select Outputs to compare
	[rc, osetSet, vecSet] = App.feSelectOutput( "Select Output Vectors",0, \
	constants.FOT_ANY, constants.FOC_ANY, constants.FT_ELEM, False, osetSet, vecSet )
	
	if False : # Change to 'if True :' to explore info below 
		## CAUTION: Set objects populated via '.feSelectOutput' do not define properly: ##
		# Return as 'win32com.client.CDispatch' instead of 'PyFemap.ISet' as expected
		# Calling methods on the 'win32com' without arguments (Such as .Count()) will
		# fail if parentheses are used. Example:
		
		# Create a regular FEMAP Set and populate with all elements
		eset = App.feSet
		eset.AddAll(constants.FT_ELEM)
		
		# Methods call normally on this Set 
		eset.Count() # works just fine
		
		# Sets from "feSelectOutput" will fail if parentheses used on methods without arguments
		vecSet.Count # will work just fine 
		# vecSet.Count() # will raise an exception
		
		# Methods that return 'win32com.client.CDispatch' objects: 
		# .NextLoad() .NextLoadDef() .feFileAttachInfo() .feSelectOutputSets()
		# .feSelectOutput() .NextBC() .NextBCDef()
		##                                                                              ##
		
	
	# Exit if valid selection not made
	if rc != -1 :
		return
	
	# Copy FEMAP Sets as Python Lists using custom function "feSetIn()"
	oSetIDs = feSetIn(osetSet)
	vecIDs = feSetIn(vecSet)
	
	# Prompt user to specify min or max values
	rc = App.feAppMessageBox( 3, "Ok to Build Table of Maximum Values (No=Minimum) ?" )
	mode = abs(rc) # set "mode" accordingly
	
	# Exit function if invalid input passed
	if mode == 2 :
		return
		
	# For each Vector ID, add all output sets to Results Browsing 
	# Object to find max/min values/IDs 
	for vID in vecIDs :
		
		# Re-initialize RBO to analyze new Vector across multiple Output Sets
		res.clear()
		
		# Initialize "vec_data" to hold all data for current Vector
		vec_data = [[], [], [], [], []] #[[minID] [maxID] [minVal] [maxVal]]
		
		# set to True when "v.Get()" command is successful
		flag = False
		osID_vect = [] # List to hold all Output Set IDs 
		
		# Analyze Vector across all chosen Output Sets
		for osID in oSetIDs :
		
			# Attempt to add to RBO 
			rc = res.AddColumn(osID,vID,False,0,[])
			
			# if Vector exists in Output Set osID
			if rc[0] == -1 :
				osID_vect.append(osID)
				
				# "Get()" Vector "vID" only once per Vector
				if not flag:
					v.GetFromSet(osID, vID)
					flag = True # short-circuit next "v.Get()" call
					
		# Ignore Nodal vectors
		if v.location == 7 :
			continue # Skips to next iteration of "for vID in vecIDs :"
		
		# Populate Results Browsing Object for analysis
		res.Populate()
		
		# Store RBO min/max column data into "vec_data"
		for i in range(res.NumberOfColumns()) :
			rc = res.GetColumnMinMax(i,0,0,0,0,0)
			
			# Store returned outputs; loop is for syntactical convenience
			for j in range(4) :
				vec_data[j].append(rc[j+1])
			
			# Store Output Set ID in "vec_Data"
			vec_data[4].append(osID_vect[i])
		
		# Find min/max values/IDs
		vec_min = min(vec_data[2])
		vec_max = max(vec_data[3])
		# The following lines function similar to a "Lookup" Excel call
		vec_minID = vec_data[0][ vec_data[2].index(vec_min) ]
		vec_maxID = vec_data[1][ vec_data[3].index(vec_max) ]
		osID_min = vec_data[4][ vec_data[2].index(vec_min) ]
		osID_max = vec_data[4][ vec_data[3].index(vec_max) ]
		# Lookup value "vec_max" in column 3, return value from column 4
		
		# Append Vector data to master data list "dataOut" 
		dataOut[3].append( [vec_min, vec_max][mode] ) # min or max based on "mode"
		dataOut[2].append( [osID_min, osID_max][mode] )
		dataOut[0].append( [vec_minID, vec_maxID][mode] )
		
		# Append Vector label to appropriate column
		dataOut[1].append(str(vID)+".."+v.title)
	
	key = False # to remember if table was locked/unlocked before API ran 
	
	# Prepare Data Table for data entry
	if table.Locked :
		key = True
		table.Lock(False)
	
	table.clear()
	# Add all data stored in "dataOut" to Data Table
	table.AddColumn(False,False,constants.FT_ELEM,0,"Vector ID",3,len(dataOut[0]),dataOut[0],dataOut[1])
	table.AddColumn(False,False,constants.FT_ELEM,0,"Output Set ID",1,len(dataOut[0]),dataOut[0],dataOut[2])
	table.AddColumn(False,False,constants.FT_ELEM,0,["Min", "Max"][mode]+" Value",2,len(dataOut[0]),dataOut[0],dataOut[3])	

	# Leave it the way you found it!
	if key : 
		table.Lock(True)
	
	# Print confirmation
	App.feAppMessage(constants.FCM_NORMAL,["Min", "Max"][mode]+" Values for "+str(len(dataOut[0]))+" Vectors across "+str(len(oSetIDs))+" Output Sets written to Data Table.")
	
	# Terminate main() Function successfully
	return		

# Invoke main function
main()

App.feAppMessage(0,"Python API Finished")