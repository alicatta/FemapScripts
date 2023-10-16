"""
This Script will convert Caesar II Neutral File into FEMAP neutral file
Created By: MA - 13/11/18 v0.1
"""

import os
import pandas as pd
import math


inputFile="C:\\Job\\Technip - Thailand\\CAP-C-02-1360-254_REV-0_FOR FIV.CII"
outputFile="C:\\Job\\Technip - Thailand\\CAP-C-02-1360-254_REV-0_FOR FIV.neu"

# Parameters:
AUTO_ADD_FLUID_MASS = True  # Will automatically add the fluid mass as NSM in Property
AUTO_SPLIT_ELEMENT = True  # Auto Subdivise Elements
LENGTH_DIAM_RATIO = 1.5  # If AUTO_SPLIT_ELEMENT, Ratio between pipe element length and diameter
MIN_ELEMENT_SIZE = 0.005  # In meter, Element smaller than this will be merged
C2RIGID_TO_NASTRANRIGID = True  # Will create a RBE2 Element if elements is defined as Rigid In Caesar
CREATE_RESTRAIN = True #Will Convert C2 restraint data
RESTRAIN_TYPE_SPRING = True #If true restrain will be type "Node To Ground" if false "SPC"
RESTRAIN_STIFFNESS = 1e7 #Node to ground Stiffness in N/m



# Material Prop:
MaterialName = 'Steel'
youngMod = 210000000000.
poisson = 0.3
densityPipe = 7850.


# data frame for differents elements:
nodes = pd.DataFrame(columns=['id', 'x', 'y', 'z', 'section'])
bends = pd.DataFrame(columns=['rad', 'tck', 'n1', 'n2'])
rigids = pd.DataFrame(columns=['mass','midNode'])
reducers = pd.DataFrame(columns=['diam', 'tck'])
restraints = pd.DataFrame(columns=['n1','n2','type'])
elements = pd.DataFrame(columns=['n1', 'n2', 'diam' , 'thick', 'bend', 'rigid', 'restraint', 'disp', 'reducer', 'density'])
properties = pd.DataFrame(columns=['name','OD','thick','density','area','I','J','perimeter','OD2','thick2','area2','I2','J2','perimeter2','tapered'])


#Datframe with Pipe Schedule:
NPS = pd.DataFrame( [[0.125,3,10.29,0.889,1.245,None,1.448,1.727,None,None,2.413,None,None,None,None,None,None],
[0.25,6,13.72,1.245,1.651,None,1.854,2.235,None,None,3.023,None,None,None,None,None,None],
[0.375,10,17.15,1.245,1.651,None,1.854,2.311,None,None,3.200,None,None,None,None,None,None],
[0.5,15,21.34,1.651,2.108,None,2.413,2.769,None,None,3.734,None,None,None,None,4.775,7.468],
[0.75,20,26.67,1.651,2.108,None,2.413,2.870,None,None,3.912,None,None,None,None,5.563,7.823],
[1,25,33.40,1.651,2.769,None,2.896,3.378,None,None,4.547,None,None,None,None,6.350,9.093],
[1.25,32,42.16,1.651,2.769,None,2.972,3.556,None,None,4.851,None,None,None,None,6.350,9.703],
[1.5,40,48.26,1.651,2.769,None,3.175,3.683,None,None,5.080,None,None,None,None,7.137,10.160],
[2,50,60.33,1.651,2.769,None,3.175,3.912,None,None,5.537,None,None,6.350,None,8.738,11.074],
[2.5,65,73.03,2.108,3.048,None,4.775,5.156,None,None,7.010,None,None,7.620,None,9.525,14.021],
[3,80,88.90,2.108,3.048,None,4.775,5.486,None,None,7.620,None,None,8.890,None,11.125,15.240],
[3.5,90,101.60,2.108,3.048,None,4.775,5.740,None,None,8.077,None,None,None,None,None,16.154],
[4,100,114.30,2.108,3.048,None,4.775,6.020,None,None,8.560,None,None,11.100,None,13.487,17.120],
[4.5,115,127.00,None,None,None,None,6.274,None,None,9.017,None,None,None,None,None,18.034],
[5,125,141.30,2.769,3.404,None,None,6.553,None,None,9.525,None,None,12.700,None,15.875,19.050],
[6,150,168.28,2.769,3.404,None,None,7.112,None,None,10.973,None,None,14.275,None,18.263,21.946],
[7,175,193.68,None,None,None,None,7.645,None,None,12.700,None,None,None,None,None,22.225],
[8,200,219.08,2.769,3.759,6.350,7.036,8.179,None,10.312,12.700,None,15.062,18.263,20.625,23.012,22.225],
[9,225,244.48,None,None,None,None,8.687,None,None,12.700,None,None,None,None,None,None],
[10,250,273.05,3.404,4.191,6.350,7.798,9.271,9.271,12.700,12.700,15.088,18.237,21.412,25.400,28.575,None],
[12,300,323.85,3.962,4.572,6.350,8.382,9.525,10.312,14.275,12.700,17.450,21.412,25.400,28.575,33.325,None],
[14,350,355.60,3.962,6.350,7.925,9.525,9.525,11.100,15.062,12.700,19.050,23.800,27.762,31.750,35.712,None],
[16,400,406.40,4.191,6.350,7.925,9.525,9.525,12.700,16.662,12.700,21.412,26.187,30.937,36.500,40.488,None],
[18,450,457.20,4.191,6.350,7.925,11.100,9.525,14.275,19.050,12.700,23.800,29.362,34.925,39.675,45.237,None],
[20,500,508.00,4.775,6.350,9.525,12.700,9.525,15.062,20.625,12.700,26.187,32.512,38.100,44.450,49.987,None],
[22,550,558.80,4.775,6.350,9.525,12.700,9.525,None,22.225,12.700,28.575,34.925,41.275,47.625,53.975,None],
[24,600,609.60,5.537,6.350,9.525,14.275,9.525,17.450,24.587,12.700,30.937,38.887,46.025,52.375,59.512,None],
[26,650,660.400,None,7.925,12.700,None,9.525,None,None,None,None,None,None,None,None,None],
[28,700,711.200,None,7.925,12.700,15.875,9.525,None,None,None,None,None,None,None,None,None],
[30,750,762.000,6.350,7.925,12.700,15.875,9.525,None,None,None,None,None,None,None,None,None],
[32,800,812.800,None,7.925,12.700,15.875,9.525,17.475,None,None,None,None,None,None,None,None],
[34,850,863.600,None,7.925,12.700,15.875,9.525,17.475,None,None,None,None,None,None,None,None],
[36,900,914.400,None,7.925,12.700,None,9.525,None,None,None,None,None,None,None,None,None]],columns=['NPS','DM','OD','Sch5s','Sch10','Sch20','Sch30','Sch40s','Sch40','Sch60','Sch80s','Sch80','Sch100','Sch120','Sch140','Sch160','XXS']).set_index('NPS')


sectionNb = 1
startPos = 0
factor_length = 0
factor_mass = 0
factor_density = 0
factor_thick = 0
factor_diameter = 0
firstNode = False

def val(line2read,pos):
    if pos in [1,2,3,4,5,6]:
        return float(line2read[2+(pos-1)*13:2+(pos)*13])
    else:
        return -1

def lineSkip(file,nb):
    for i in range(nb):
        file.readline()

def vlen(v):
    lenght = 0
    for c in v:
        lenght += c**2
    return math.sqrt(lenght)

def inputoutput():
    fileNotFound = True
    #input file name
    while fileNotFound:
        inputFile = input("Enter Caesar II Neutral File (*.cii) to read: ")
        if inputFile[-4:].lower() != ".cii":
            inputFile = inputFile + ".cii"
        if os.path.isfile(inputFile.lower()):
            fileNotFound = False
        else:
            print('Couldn\'t find file: '+ inputFile.lower())

    #output file name
    fileNotValid = True
    while fileNotValid:
        outputFile = input("Enter FEMAP Neutral File (*.neu) to write: ")
        if outputFile[-4:].lower() != ".neu":
            outputFile = outputFile + ".neu"
        if os.path.isfile(outputFile.lower()):
            print('File: ' + outputFile.lower() + ' already exist...')
        else:
            fileNotValid = False

    return (inputFile,outputFile)

def readUnitSys(f):
    global factor_length,factor_mass,factor_density,factor_thick,factor_diameter
    print('Reading Unit System...', end='')
    line =f.readline()
    #reading Unit Header
    while f.readline().strip() != "#$ UNITS":
        pass
    line = f.readline()
    factor_length = 25.4 / val(line,1)/1000
    factor_mass = 0.1019 * 4.448 / val(line,2)
    lineSkip(f, 1)
    line = f.readline()
    factor_density = 27680 / val(line,1)
    line = f.readline()
    factor_thick = 25.4 / val(line,4)/1000
    factor_diameter = 25.4 / val(line,3)/1000
    lineSkip(f, 2)
    lon = f.readline().strip()
    Force = f.readline().strip()
    lineSkip(f, 10)
    dens = f.readline().strip()
    lineSkip(f, 7)
    diam = f.readline().strip()
    thick = f.readline().strip()
    print('Done')

    print("Caesar Unit System:")
    print("Lenght: "+ lon + " Diameter: " + diam + " Thickness: "+ thick + " Weight: " + Force + " Density " + dens)
    print("Export to FEMAP is made in SI units")

def readCoords(f):
    global sectionNb,nodes,firstNode
    # reading COORDS Header
    f.seek(startPos)
    print('Reading Coordinates...', end='')
    CoordsExist = False
    line = f.readline()
    while line:
        if line.strip() == "#$ COORDS":
            CoordsExist = True
            break
        line = f.readline()
    if CoordsExist:
        nbCoord = int(f.readline())
        for i in range(nbCoord):
            line = f.readline()
            nodes = nodes.append({'id': val(line,1),
                          'x': val(line, 2)*factor_length,
                          'y': val(line, 3)*factor_length,
                          'z': val(line, 4)*factor_length,
                          'section': sectionNb},ignore_index=True)
            sectionNb += 1

    else:
        firstNode = True
    print('Done')

def readBends(f):
    # reading Bend Header
    global bends
    print('Reading Bend...', end='')
    f.seek(startPos)
    while f.readline().strip() != "#$ BEND":
        pass

    line = f.readline()
    while line.strip() != "#$ RIGID":
        bends = bends.append({'rad': val(line, 1)*factor_length,
                              'tck': val(f.readline(), 4)*factor_thick,
                              'n1': val(line, 4),
                              'n2': val(line, 6), }, ignore_index=True)
        line = f.readline()
        line = f.readline()
    print('Done')

def readRigids(f):
    # reading Rigid Header
    global rigids
    print('Reading Rigid...', end='')
    f.seek(startPos)
    while f.readline().strip() != "#$ RIGID":
        pass
    line = f.readline()
    while line.strip() != "#$ EXPJT":
        rigids = rigids.append({'mass': val(line, 1)*factor_mass}, ignore_index=True)
        line = f.readline()
    print('Done')

def readReducers(f):
    # reading Reducer Header
    global reducers
    print('Reading Reducers...', end='')
    f.seek(startPos)
    while f.readline().strip() != "#$ REDUCERS":
        pass
    line = f.readline()
    while line.strip() != "#$ FLANGES":
        reducers = reducers.append({'diam': val(line, 1)*factor_diameter,
                              'tck': val(line, 2)*factor_thick,}, ignore_index=True)
        line = f.readline()
    print('Done')

def readRestraints(f):
    # reading Restraints Header
    print('Reading Restraints...', end='')
    f.seek(startPos)
    global restraints
    while f.readline().strip() != "#$ RESTRANT":
        pass
    line = f.readline()
    while line.strip() != "#$ DISPLMNT":
        if val(line,1) > 1:
            restraints = restraints.append({'n1': val(line, 1),
                                  'n2': val(line, 6),
                                  'type': val(line, 2),}, ignore_index=True)
        line = f.readline()
    print('Done')

def readElements(f):
    # reading Elements Header
    print('Reading Elements...', end='')
    global elements, sectionNb, nodes, firstNode, restraints
    f.seek(startPos)
    while f.readline().strip() != "#$ ELEMENTS":
        pass
    line = f.readline()
    while line.strip() != "#$ AUX_DATA":
        n1 = val(line, 1)
        n2 = val(line, 2)
        Dx = val(line, 3) * factor_length
        Dy = val(line, 4) * factor_length
        Dz = val(line, 5) * factor_length
        diameter = val(line, 6) * factor_diameter
        line = f.readline()
        thickness = val(line, 1) * factor_thick
        lineSkip(f,3)
        fluid_dens = val(f.readline(), 2) * factor_density
        lineSkip(f, 3) #some C2 version number have to be 3 some other 4
        line = f.readline()
        bend = val(line, 1)
        rigid = val(line, 2)
        restraint = val(line,4)
        disp = val(line,5)
        line = f.readline()
        tee = val(line,5)
        line = f.readline()
        reducer = val(line,1)
        line = f.readline()
        if firstNode:
            nodes = nodes.append({'id': n1,'x': 0.,'y': 0.,'z': 0.,'section':sectionNb }, ignore_index=True)
            sectionNb += 1
            firstNode = False
        if nodes.loc[nodes['id'] == n1].shape[0] > 0: #looking if N1 already exist
            node1 = nodes.loc[nodes['id'] == n1].iloc[0]
        else:
            if restraints.loc[restraints['n2'] == n1].shape[0] > 0:
                node1 = nodes.loc[nodes['id'] == restraints.loc[restraints['n2'] == n1].iloc[0]['n1']].iloc[0]
            else:
                nodes = nodes.append({'id': n1, 'x': 0., 'y': 0., 'z': 0., 'section': sectionNb}, ignore_index=True)
                sectionNb += 1
                node1 = nodes.loc[nodes['id'] == n1].iloc[0]

        addingN2 = True
        replacingN2 = False
        if nodes.loc[(nodes['id'] == n2) & (nodes['section'] == node1.section)].shape[0] > 0:
            addingN2 = False
            if tee == 0:
                replacingN2 = True

        if addingN2:
            nodes = nodes.append({'id': n2, 'x': node1.x + Dx, 'y': node1.y + Dy, 'z': node1.z + Dz, 'section': node1.section }, ignore_index=True)

        if replacingN2:
            nodes.loc[(nodes['id'] == n2) & (nodes['section'] == node1.section),'x'] = node1.x + Dx
            nodes.loc[(nodes['id'] == n2) & (nodes['section'] == node1.section),'y'] = node1.y + Dy
            nodes.loc[(nodes['id'] == n2) & (nodes['section'] == node1.section),'z'] = node1.z + Dz

        
        elements = elements.append({'n1': node1.id,
                                            'n2': n2,
                                            'diam': diameter,
                                            'thick': thickness,
                                            'bend':bend,
                                            'rigid':rigid,
                                            'restraint':restraint,
                                            'disp':disp,
                                            'reducer':reducer,
                                            'density':fluid_dens }, ignore_index=True)
    print('Done')

def sectionAssembly():
    global nodes
    print('Assembling Sections...', end='')
    duplicates = nodes.loc[nodes.duplicated('id',False)]
    if duplicates.section.unique().__len__() != sectionNb -1:
        print('/!\ Warning: Some sections don\'t connect with each other.')
    while duplicates.shape[0] != 0:
        NodeToMerge = duplicates.loc[duplicates['section'] != 1].iloc[0]
        NodeOriginal = duplicates.loc[(duplicates['id']==NodeToMerge['id'])&(duplicates['section']!=NodeToMerge['section'])].iloc[0]
        Dx = NodeOriginal['x']-NodeToMerge['x']
        Dy = NodeOriginal['y'] - NodeToMerge['y']
        Dz = NodeOriginal['z'] - NodeToMerge['z']
        nodes = nodes.drop(nodes.loc[(nodes['id'] == NodeToMerge.id) & (nodes['section'] == NodeToMerge.section)].index)
        sameSectionNodes = duplicates.loc[(duplicates['id']!=NodeToMerge['id'])&(duplicates['section']==NodeToMerge['section'])]
        for index, row in sameSectionNodes.iterrows():
            nodes = nodes.drop(index)
        nodes.loc[nodes['section'] == NodeToMerge.section, 'x'] = nodes.loc[nodes['section'] == NodeToMerge.section, 'x'] + Dx
        nodes.loc[nodes['section'] == NodeToMerge.section, 'y'] = nodes.loc[nodes['section'] == NodeToMerge.section, 'y'] + Dy
        nodes.loc[nodes['section'] == NodeToMerge.section, 'z'] = nodes.loc[nodes['section'] == NodeToMerge.section, 'z'] + Dz
        nodes.loc[nodes['section'] == NodeToMerge.section, 'section'] = NodeOriginal.section
        duplicates = nodes.loc[nodes.duplicated('id', False)]
    print('Done')

def createBends():
    global elements, nodes
    print('Converting Bends...', end='')
    for index, row in elements.iterrows():
        row = elements.loc[index]
        if row.bend > 0:
            bendRadius= bends.iloc[int(row.bend)-1].rad
            bendThickness = bends.iloc[int(row.bend)-1].tck
            ie1 = index
            e1 = row
            in1 = nodes.loc[nodes['id']== row.n1].index[0]
            n1 = nodes.loc[nodes['id']== row.n1].iloc[0]
            in2 = nodes.loc[nodes['id'] == row.n2].index[0]
            n2 = nodes.loc[nodes['id'] == row.n2].iloc[0]
            ie2 = elements.loc[elements['n1'] == n2.id].index[0]
            e2 = elements.loc[elements['n1'] == n2.id].iloc[0]
            in3 = nodes.loc[nodes['id'] == e2.n2].index[0]
            n3 = nodes.loc[nodes['id'] == e2.n2].iloc[0]
            if bendRadius <= 0:
                bendRadius = e1.diam * 1
            if bendThickness <= 0:
                bendThickness = e1.thick
            dx = n2.x - n1.x - bendRadius

            angle_v1 = [n1.x - n2.x, n1.y - n2.y, n1.z - n2.z]
            angle_v2 = [n3.x - n2.x, n3.y - n2.y, n3.z - n2.z]

            cos_angle = (angle_v1[0] * angle_v2[0] + angle_v1[1] * angle_v2[1] + angle_v1[2] * angle_v2[2]) / (vlen(angle_v1) * vlen(angle_v2))
            angle = math.acos(cos_angle)

            idMax = nodes['id'].max()

            if bends.iloc[int(row.bend)-1].n1>0:
                id4 = bends.iloc[int(row.bend)-1].n1
            else:
                idMax += 10
                id4 = idMax

            if bends.iloc[int(row.bend)-1].n2>0:
                id5 = bends.iloc[int(row.bend)-1].n2
            else:
                idMax += 10
                id5 = idMax

            n4 = pd.Series({'x': n2.x - (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v1)) * (n2.x - n1.x),
                            'y': n2.y - (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v1)) * (n2.y - n1.y),
                            'z': n2.z - (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v1)) * (n2.z - n1.z),
                            'id':id4})

            n5 = pd.Series({'x': n2.x - (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v2)) * (n2.x - n3.x),
                            'y': n2.y - (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v2)) * (n2.y - n3.y),
                            'z': n2.z - (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v2)) * (n2.z - n3.z),
                            'id': id5})
            if (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v1)) > 1:
                print("Error: Node {} Bend First Node too short for bend radius".format(n2.id))
            else:
                if (bendRadius * math.tan((3.14159 - angle) / 2) / vlen(angle_v2)) > 1:
                    print("Error: Node {} Bend Second Node too short for bend radius".format(n2.id))
                else:


                    v1 = [n1.x - n4.x, n1.y - n4.y, n1.z - n4.z]
                    v2 = [n5.x - n4.x, n5.y - n4.y, n5.z - n4.z]

                    vPlan = [v1[1] * v2[2] - v1[2] * v2[1],v1[2] * v2[0] - v1[0] * v2[2],v1[0] * v2[1] - v1[1] * v2[0]]
                    VPerp = [vPlan[1] * v1[2] - vPlan[2] * v1[1],vPlan[2] * v1[0] - vPlan[0] * v1[2],vPlan[0] * v1[1] - vPlan[1] * v1[0]]

                    VPerpL = vlen(VPerp)

                    VPerp = [x/VPerpL * bendRadius for x in VPerp]
                    n0 = pd.Series({'x': n4.x+ VPerp[0],
                                    'y': n4.y+ VPerp[1],
                                    'z': n4.z+ VPerp[2]})


                    n6 = pd.Series({'x': n4.x+ (n5.x-n4.x)*0.25,
                                    'y': n4.y+ (n5.y-n4.y)*0.25,
                                    'z': n4.z+ (n5.z-n4.z)*0.25,
                                    'id': idMax+10})
                    n7 = pd.Series({'x': n4.x + (n5.x - n4.x) * 0.5,
                                    'y': n4.y + (n5.y - n4.y) * 0.5,
                                    'z': n4.z + (n5.z - n4.z) * 0.5,
                                    'id': idMax + 20})
                    n8 = pd.Series({'x': n4.x + (n5.x - n4.x) * 0.75,
                                    'y': n4.y + (n5.y - n4.y) * 0.75,
                                    'z': n4.z + (n5.z - n4.z) * 0.75,
                                    'id': idMax + 30})
                    idMax += 30

                    NVertors = []
                    NVertors.append([n6.x - n0.x, n6.y - n0.y, n6.z - n0.z])
                    NVertors.append([n7.x - n0.x, n7.y - n0.y, n7.z - n0.z])
                    NVertors.append([n8.x - n0.x, n8.y - n0.y, n8.z - n0.z])

                    for i in range(3):
                        vlenght = vlen(NVertors[i])
                        NVertors[i] = [x / vlenght * bendRadius for x in NVertors[i]]

                    n6.x = n0.x + NVertors[0][0]
                    n6.y = n0.y + NVertors[0][1]
                    n6.z = n0.z + NVertors[0][2]
                    n7.x = n0.x + NVertors[1][0]
                    n7.y = n0.y + NVertors[1][1]
                    n7.z = n0.z + NVertors[1][2]
                    n8.x = n0.x + NVertors[2][0]
                    n8.y = n0.y + NVertors[2][1]
                    n8.z = n0.z + NVertors[2][2]

                    nodes = nodes.append(n6, ignore_index=True)
                    nodes = nodes.append(n7, ignore_index=True)
                    nodes = nodes.append(n8, ignore_index=True)
                    elements = elements.append({'n1': n6.id,
                                                'n2': n7.id,
                                                'diam': e1.diam,
                                                'thick': bendThickness,
                                                'bend': 0,
                                                'rigid': 0,
                                                'restraint': 0,
                                                'disp': 0,
                                                'reducer': 0,
                                                'density': e1.density}, ignore_index=True)
                    elements = elements.append({'n1': n7.id,
                                                'n2': n8.id,
                                                'diam': e1.diam,
                                                'thick': bendThickness,
                                                'bend': 0,
                                                'rigid': 0,
                                                'restraint': 0,
                                                'disp': 0,
                                                'reducer': 0,
                                                'density': e1.density}, ignore_index=True)
                    if vlen([n1.x-n4.x,n1.y-n4.y,n1.z-n4.z])>MIN_ELEMENT_SIZE:
                        nodes = nodes.append(n4, ignore_index=True)
                        elements.loc[ie1, 'n2'] = id4
                        elements = elements.append({'n1': n4.id,
                                                    'n2': n6.id,
                                                    'diam': e1.diam,
                                                    'thick': bendThickness,
                                                    'bend': 0,
                                                    'rigid': 0,
                                                    'restraint': 0,
                                                    'disp': 0,
                                                    'reducer': 0,
                                                    'density': e1.density}, ignore_index=True)
                    else:
                        elements.loc[ie1, 'n2'] = n6.id
                    if vlen([n3.x - n5.x, n3.y - n5.y, n3.z - n5.z]) > MIN_ELEMENT_SIZE:
                        nodes = nodes.append(n5, ignore_index=True)
                        elements.loc[ie2, 'n1'] = id5
                        elements = elements.append({'n1': n8.id,
                                                    'n2': n5.id,
                                                    'diam': e1.diam,
                                                    'thick': bendThickness,
                                                    'bend': 0,
                                                    'rigid': 0,
                                                    'restraint': 0,
                                                    'disp': 0,
                                                    'reducer': 0,
                                                    'density': e1.density}, ignore_index=True)
                    else:
                        elements.loc[ie2, 'n1'] = n8.id
                    nodes = nodes.drop(
                        nodes.loc[(nodes['id'] == n2.id)].index)
    print('Done')

def removeSmallElements():
    global elements, nodes
    print('Removing Small Elements...', end='')
    for index, row in elements.iterrows():
        row = elements.loc[index]
        n1 = nodes.loc[nodes['id'] == row.n1].iloc[0]
        n2 = nodes.loc[nodes['id'] == row.n2].iloc[0]

        if vlen([n1.x - n2.x, n1.y - n2.y, n1.z - n2.z]) < MIN_ELEMENT_SIZE:
            nodeToRemove = n1 if n1.id>n2.id else n2
            nodeToKeep = n2 if n1.id>n2.id else n1
            elemsToModify = elements.loc[(elements['n1'] == nodeToRemove.id) & (elements.index != index)].index
            for indexElem in elemsToModify:
                elements.loc[indexElem,'n1']=nodeToKeep['id']
            elemsToModify = elements.loc[(elements['n2'] == nodeToRemove.id) & (elements.index != index)].index
            for indexElem in elemsToModify:
                elements.loc[indexElem, 'n2'] = nodeToKeep['id']
            if nodeToRemove.id != nodeToKeep.id:
                nodes = nodes.drop(nodes.loc[(nodes['id'] == nodeToRemove.id)].index)
            elements = elements.drop(index)



def defineProperties():
    global elements,properties
    print('Creating Properties...', end='')
    for index, row in elements.iterrows():
        diam = round(row.diam,5)
        thick = round(row.thick,5)
        density = round(row.density,1)
        diamI = diam - 2*thick
        if row.rigid != 0 & C2RIGID_TO_NASTRANRIGID:
            pass
        else:
            if row.reducer == 0:
                if not properties.loc[(properties['OD']==diam)&(properties['thick']==thick)&(properties['tapered']==False)&((properties['density']==density) | (not AUTO_ADD_FLUID_MASS))].empty:
                    elements.loc[index,"property"] = properties.loc[(properties['OD']==diam)&(properties['thick']==thick)&(properties['tapered']==False)&((properties['density']==density) | (not AUTO_ADD_FLUID_MASS))].index[0]
                else:
                    properties = properties.append({'name':'',
                                                    'OD':diam,
                                                    'thick':thick,
                                                    'density':density,
                                                    'area':math.pi * (diam**2-diamI**2)/4,
                                                    'I':math.pi*(diam**4-diamI**4)/64,
                                                    'J':math.pi*(diam**4-diamI**4)/32,
                                                    'perimeter':math.pi*diam,
                                                    'tapered': False}, ignore_index=True)
                    elements.loc[index, "property"] = properties.index[-1]
            else:
                diam2 = round(reducers.iloc[int(row.reducer)-1].diam, 5)
                thick2 = round(reducers.iloc[int(row.reducer)-1].tck, 5)
                diamI2 = diam2 - 2 * thick2

                if not properties.loc[(properties['OD']==diam)&(properties['thick']==thick)&((properties['density']==density) | (not AUTO_ADD_FLUID_MASS)) & (properties['thick2']==thick2) & (properties['OD2']==diam2)].empty:
                    elements.loc[index,"property"] = properties.loc[(properties['OD']==diam)&(properties['thick']==thick)&((properties['density']==density) | (not AUTO_ADD_FLUID_MASS)) & (properties['thick2']==thick2) & (properties['OD2']==diam2)].index[0]
                else:
                    properties = properties.append({'name':'',
                                                    'OD':diam,
                                                    'thick':thick,
                                                    'density':density,
                                                    'area':math.pi * (diam**2-diamI**2)/4,
                                                    'I':math.pi*(diam**4-diamI**4)/64,
                                                    'J':math.pi*(diam**4-diamI**4)/32,
                                                    'perimeter':math.pi*diam,
                                                    'OD2':diam2,
                                                    'thick2':thick2,
                                                    'area2':math.pi * (diam2**2-diamI2**2)/4,
                                                    'I2':math.pi*(diam2**4-diamI2**4)/64,
                                                    'J2':math.pi*(diam2**4-diamI2**4)/32,
                                                    'perimeter2':math.pi*diam2,
                                                    'tapered': True}, ignore_index=True)
                    elements.loc[index, "property"] = properties.index[-1]





    for index in properties.index:
        size = '{:.2f}mm'.format(properties.loc[index,'OD']*1000)
        sch =  '{:.2f}mm'.format(properties.loc[index,'thick']*1000)
        red = ''
        if not NPS.loc[round(NPS['OD']/1000,4)==round(properties.loc[index,'OD'],4)].empty:
            size = str(NPS.loc[round(NPS['OD']/1000,4) == round(properties.loc[index, 'OD'], 4)].index[0]) + "''"
            series = NPS.loc[round(NPS['OD']/1000,4) == round(properties.loc[index, 'OD'], 4)].iloc[0]
            if not series[round(series/1000,4) == round(properties.loc[index, 'thick'],4)].empty:
                sch = series[round(series/1000,4) == round(properties.loc[index, 'thick'],4)].index[0]
        if properties.loc[index,'tapered']:
            red='_reducer'
        properties.loc[index,'name']= 'pipe_'+size+'_'+sch+red

    print('Done')

def rigidMidNode():
    global elements, nodes
    idMax = nodes['id'].max()
    for index, row in elements.iterrows():
        if row.rigid != 0:
            n1 = nodes.loc[nodes['id']== row.n1].iloc[0]
            n2 = nodes.loc[nodes['id'] == row.n2].iloc[0]

            n3 = pd.Series({'x': (n1.x+n2.x)/2,
                            'y': (n1.y+n2.y)/2,
                            'z': (n1.z+n2.z)/2,
                            'id': idMax + 10})
            idMax= idMax +10
            nodes = nodes.append(n3, ignore_index=True)
            rigids.loc[int(row.rigid-1),"midNode"] = n3['id']


def writeHeader(f):
    print('        File Header...', end='')
    f.write('   -1\n')
    f.write('   100\n')
    f.write('<NULL>\n')
    f.write('11.4,\n')
    f.write('   -1\n')
    f.write('   -1\n')
    f.write('   405\n')
    f.write('   -1\n')
    f.write('   -1\n')
    f.write('   475\n')
    f.write('   -1\n')
    f.write('   -1\n')
    f.write('   410\n')
    f.write('   -1\n')
    f.write('   -1\n')
    f.write('  1150\n')
    f.write('   -1\n')
    f.write('   -1\n')
    f.write('  1151\n')
    f.write('   -1\n')
    f.write('   -1\n')
    f.write('   413\n')
    f.write('1,124,\n')
    f.write('Default Layer\n')
    f.write('9999,124,\n')
    f.write('Construction Layer\n')
    f.write('   -1\n')
    print('Done')

def writeMaterial(f):
    print('        Material...', end='')
    f.write('   -1\n')
    f.write('   601\n')
    f.write('1,-601,55,0,0,1,0\n')
    f.write("{:.79}\n".format(MaterialName))
    f.write('10,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('25,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,\n')
    f.write('200,\n')
    f.write('{},{},{},0.,0.,0.,{},{},{},0.,\n'.format(youngMod,youngMod,youngMod,poisson,poisson,poisson))
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,{},\n'.format(densityPipe))
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('0.,0.,0.,0.,0.,0.,0.,0.,0.,0.,\n')
    f.write('50,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('70,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('0,0,0,0,0,0,0,0,0,0,\n')
    f.write('   -1\n')
    print('Done')

def writeProperties(f):
    print('        Properties...', end='')
    f.write('   -1\n')
    f.write('   402\n')
    for index, row in properties.iterrows():
        f.write("{},110,1,5,1,0,0,\n".format(index+1))
        f.write("{}\n".format(row['name']))
        if (row.tapered):
            f.write('1,6,1,1,\n')
        else:
            f.write('0,6,1,0,\n')
        f.write('10,\n')
        f.write('0,0,0,0,0,0,0,0,\n')
        f.write('0,0,\n')
        f.write('78,\n')
        f.write('{:.8f},{:.10f},{:.10f},0.,{:.10f},\n'.format(row.area,row['I'],row['I'],row['J']))
        f.write('0.,0.,{:.5f},0.,-{:.5f},\n'.format((row.density*row.area if AUTO_ADD_FLUID_MASS else 0.),row.OD/2))
        f.write('{:.5f},0.,0.,{:.5f},-{:.5f},\n'.format(row.OD/2,row.OD/2,row.OD/2))
        f.write('0.,0.,0.,0.,{:.6f},\n'.format(row.perimeter))
        if (row.tapered):
            f.write('{:.8f},{:.10f},{:.10f},0.,{:.10f},\n'.format(row.area2, row['I2'], row['I2'], row['J2']))
            f.write('0.,0.,{:.5f},0.,-{:.5f},\n'.format((row.density * row.area2 if AUTO_ADD_FLUID_MASS else 0.),row.OD2 / 2))
            f.write('{:.5f},0.,0.,{:.5f},-{:.5f},\n'.format(row.OD / 2, row.OD2 / 2, row.OD2 / 2))
            f.write('0.,0.,0.,0.,{:.6f},\n'.format(row.perimeter2))
        else:
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
        f.write('{:.5f},0.,0.,0.,0.,\n'.format(row.OD/2))
        f.write('{:.5f},1.,2.,3.,4.,\n'.format(row.thick))
        f.write('0.,-1.,0.,0.,0.,\n')
        if (row.tapered):
            f.write('{:.5f},0.,0.,0.,0.,\n'.format(row.OD2 / 2))
            f.write('{:.5f},1.,2.,3.,4.,\n'.format(row.thick2))
            f.write('0.,-1.,0.,0.,0.,\n')
        else:
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
        f.write('0.3,0.,0.,0.,0.,\n')
        f.write('0.,0.,0.,\n')
        f.write('0,\n')
        f.write('0,\n')
    if C2RIGID_TO_NASTRANRIGID:
        for index,row in rigids.iterrows():
            f.write('{},110,0,27,1,0,0,\n'.format(len(properties.index)+index+1))
            f.write('MASS Rigid {},\n'.format(index+1))
            f.write('0,0,0,0,\n')
            f.write('10,\n')
            f.write('0,0,0,0,0,0,0,0,\n')
            f.write('0,0,\n')
            f.write('78,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,{:.3f},0.,0.,\n'.format(row.mass))
            f.write('0.,{:.3f},{:.3f},0.,0.,\n'.format(row.mass,row.mass))
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,0.,0.,\n')
            f.write('0.,0.,0.,\n')
            f.write('0,\n')
            f.write('0,\n')
    f.write('   -1\n')
    print('Done')

def writeNodes(f):
    print('        Nodes...', end='')
    f.write('   -1\n')
    f.write('   403\n')
    for index, row in nodes.iterrows():
        f.write("{},0,0,1,46,0,0,0,0,0,0,{:.4f},{:.4f},{:.4f},0,0,\n".format(int(row.id),row.x,row.y,row.z))
    f.write('   -1\n')
    print('Done')

def writeElems(f):
    extraId = len(elements.index)+100
    print('        Elements...', end='')
    f.write('   -1\n')
    f.write('   404\n')
    for index, row in elements.iterrows():
        if row.rigid != 0 and C2RIGID_TO_NASTRANRIGID:
            midNode = rigids.loc[int(row.rigid-1),"midNode"]
            massProp = len(properties.index)+int(row.rigid)
            f.write("{},124,0,29,13,1,0,0,0,0,0,0,0,0,\n".format(extraId))
            f.write("{},0,0,0,0,0,0,0,0,0,\n".format(int(row.n1)))
            f.write("0,0,0,0,0,0,0,0,0,0,\n")
            f.write("0.,0.,0.,0,0,0,0,0,0,\n")
            f.write("0.,0.,0.,\n")
            f.write("0.,0.,0.,\n")
            f.write("0,1,1,1,1,1,1,0,0,0,0,0,0,1,0,0,0,\n")
            f.write("{},0,0.,0,0,0,0,0,0,\n".format(int(midNode)))
            f.write("-1,\n")
            f.write("{},124,0,29,13,1,0,0,0,0,0,0,0,0,\n".format(extraId+1))
            f.write("{},0,0,0,0,0,0,0,0,0,\n".format(int(midNode)))
            f.write("0,0,0,0,0,0,0,0,0,0,\n")
            f.write("0.,0.,0.,0,0,0,0,0,0,\n")
            f.write("0.,0.,0.,\n")
            f.write("0.,0.,0.,\n")
            f.write("0,1,1,1,1,1,1,0,0,0,0,0,0,1,0,0,0,\n")
            f.write("{},0,0.,0,0,0,0,0,0,\n".format(int(row.n2)))
            f.write("-1,\n")
            f.write("{},124,{},27,9,1,0,0,0,0,0,0,0,0,\n".format(index+1,massProp))
            f.write("{},0,0,0,0,0,0,0,0,0,\n".format(int(midNode)))
            f.write("0,0,0,0,0,0,0,0,0,0,\n")
            f.write("0.,0.,0.,0,0,0,0,0,0,\n")
            f.write("0.,0.,0.,\n")
            f.write("0.,0.,0.,\n")
            f.write("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,\n")
            extraId = extraId+2
        else:
            f.write("{},124,{},5,0,1,0,0,0,0,0,0,0,0,\n".format(index+1,int(row.property+1)))
            f.write("{},{},0,0,0,0,0,0,0,0,\n".format(int(row.n1),int(row.n2)))
            f.write("0,0,0,0,0,0,0,0,0,0,\n")
            n1 = nodes.loc[nodes['id']==row.n1].iloc[0]
            n2 = nodes.loc[nodes['id']== row.n2].iloc[0]
            if (n1.x != n2.x) or (n1.y != n2.y):
                f.write("0.,0.,1.,0,0,0,0,0,0,\n")
            else:
                f.write("0.,1.,0.,0,0,0,0,0,0,\n")
            f.write("0.,0.,0.,\n")
            f.write("0.,0.,0.,\n")
            f.write("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,\n")

    f.write('   -1\n')
    print('Done')

if __name__ == '__main__':
    print("=================================================================")
    print()
    print("Caesar II Neutral File Converter to FEMAP v0.1 ---- Wood VDN 2018")
    print()
    print("=================================================================")
    print()
    #opening input file
    print('Opening Input File...', end='')
    with open(inputFile) as f:
        print('Done')
        startPos = f.tell()
        readUnitSys(f)
        readCoords(f)
        readBends(f)
        readRigids(f)
        readReducers(f)
        readRestraints(f)
        readElements(f)

    sectionAssembly()
    createBends()
    removeSmallElements()
    defineProperties()
    if C2RIGID_TO_NASTRANRIGID:
        rigidMidNode()


    with open(outFile, 'w') as f:
        print('Writting FEMAP file...')
        writeHeader(f)
        writeMaterial(f)
        writeProperties(f)
        writeNodes(f)
        writeElems(f)