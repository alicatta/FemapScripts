import pythoncom
import PyFemap
from PyFemap import constants #Sets femap constants to constants object so calling constants in constants.FCL_BLACK, not PyFemap.constants.FCL_BLACK
import sys

import win32com.client as win32
win32.gencache.is_readonly=False
appExcel = win32.Dispatch('Excel.Application') #Calls and creates new excel application

try:
    existObj = pythoncom.connect(PyFemap.model.CLSID) #Grabs active model
    app = PyFemap.model(existObj)
except:
    sys.exit("femap not open") #Exits program if there is no active femap model

rc = app.feAppMessage(0,"Python API Example Started")

workBook = appExcel.Workbooks.Add()
workSheet = workBook.Worksheets(1)

workSheet.Cells(1,1).Value = "Element Type"
workSheet.Cells(1,2).Value = "Element Count"

col = 2

pElemSet = app.feSet

Msg = "Model Element Summary"

app.feAppMessage(constants.FCL_BLACK, Msg)

for i in range(1,42):
    pElemSet.clear()
    pElemSet.AddRule(i,constants.FGD_ELEM_BYTYPE)

    if pElemSet.Count() > 0:
        workSheet.Cells(col, 2).Value = str(pElemSet.Count())

        if i == 1:
            workSheet.Cells(col,1).Value = "L_ROD elements"
        elif i == 2:
            workSheet.Cells(col, 1).Value = " L_BAR elements"
        elif i == 3:
            workSheet.Cells(col, 1).Value = " L_TUBE elements"
        elif i == 4:
            workSheet.Cells(col, 1).Value = " L_LINK elements"
        elif i == 5:
            workSheet.Cells(col, 1).Value = " L_BEAM elements"
        elif i == 6:
            workSheet.Cells(col, 1).Value = " L_SPRING elements"
        elif i == 7:
            workSheet.Cells(col, 1).Value = " L_DOF_SPRING elements"
        elif i == 8:
            workSheet.Cells(col, 1).Value = " L_CURVED_BEAM elements"
        elif i == 9:
            workSheet.Cells(col, 1).Value = " L_GAP elements"
        elif i == 10:
            workSheet.Cells(col, 1).Value = " L_PLOT elements"
        elif i == 11:
            workSheet.Cells(col, 1).Value = " L_SHEAR elements"
        elif i == 12:
            workSheet.Cells(col, 1).Value = " P_SHEAR elements"
        elif i == 13:
            workSheet.Cells(col, 1).Value = " L_MEMBRANE elements"
        elif i == 14:
            workSheet.Cells(col, 1).Value = " P_MEMBRANE elements"
        elif i == 15:
            workSheet.Cells(col, 1).Value = " L_BENDING elements"
        elif i == 16:
            workSheet.Cells(col, 1).Value = " P_BENDING elements"
        elif i == 17:
            workSheet.Cells(col, 1).Value = " L_PLATE elements"
        elif i == 18:
            workSheet.Cells(col, 1).Value = " P_PLATE elements"
        elif i == 19:
            workSheet.Cells(col, 1).Value = " L_PLANE_STRAIN elements"
        elif i == 20:
            workSheet.Cells(col, 1).Value = " P_PLANE_STRAIN elements"
        elif i == 21:
            workSheet.Cells(col, 1).Value = " L_LAMINATE_PLATE elements"
        elif i == 22:
            workSheet.Cells(col, 1).Value = " P_LAMINATE_PLATE elements"
        elif i == 23:
            workSheet.Cells(col, 1).Value = " L_AXISYM elements"
        elif i == 24:
            workSheet.Cells(col, 1).Value = " P_AXISYM elements"
        elif i == 25:
            workSheet.Cells(col, 1).Value = " L_SOLID elements"
        elif i == 26:
            workSheet.Cells(col, 1).Value = " P_SOLID elements"
        elif i == 27:
            workSheet.Cells(col, 1).Value = " L_MASS elements"
        elif i == 28:
            workSheet.Cells(col, 1).Value = " L_MASS_MATRIX elements"
        elif i == 29:
            workSheet.Cells(col, 1).Value = " L_RIGID elements"
        elif i == 30:
            workSheet.Cells(col, 1).Value = " L_STIFF_MATRIX elements"
        elif i == 31:
            workSheet.Cells(col, 1).Value = " L_CURVED_TUBE elements"
        elif  i == 32:
            workSheet.Cells(col, 1).Value = " L_PLOT_PLATE elements"
        elif i == 33:
            workSheet.Cells(col, 1).Value = " L_SLIDE_LINE elements"
        elif i == 34:
            workSheet.Cells(col, 1).Value = " L_CONTACT elements"
        elif i == 35:
            workSheet.Cells(col, 1).Value = " L_AXISYM_SHELL elements"
        elif i == 36:
            workSheet.Cells(col, 1).Value = " P_AXISYM_SHELL elements"
        elif i == 37:
            workSheet.Cells(col, 1).Value = " P_BEAM elements"
        elif i == 38:
            workSheet.Cells(col, 1).Value = " L_WELD elements"
        elif i == 39:
            workSheet.Cells(col, 1).Value = " L_SOLID_LAMINATE elements"
        elif i == 40:
            workSheet.Cells(col, 1).Value = " P_SOLID_LAMINATE elements"
        elif i == 41:
            workSheet.Cells(col, 1).Value = " L_SPRING_to_GROUND elements"
        elif i == 42:
            workSheet.Cells(col, 1).Value = " L_DOF_SRPING_to_GROUND elements"
        col = col + 1

appExcel.Visible = True
rc = app.feAppMessage(0,"Python API Example Finished")