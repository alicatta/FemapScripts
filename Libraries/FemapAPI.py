# -*- coding: utf-8 -*-

"""
WOOD VDN Femap API


RELEASE HISTORY
    v1: initial release - Martin
    v2: added function to interface with node and properties: getNodesInfoDialog getNodeInfo selectProperty
    v3: added new functions. Improved some commenting. Added table of contents (sorted by method operation) to beginning. - Alastair

PyFemap is version specific, need to be generated once calling the following script:
    import sys
    from win32com.client import makepy
    sys.argv = ["makepy", "-o Pyfemap.py", r"C:\Program Files\Siemens\Femap XXXX\femap.tlb"] 
    makepy.main() 
    
# # Table of Contents

## Create
- createNode
- createStandardMaterial
- createIsotropicMaterial
- createLumpedMassProperty
- createLumpedMass
- createBeamPipeSection
- createBendB31Property
- createBeamPropertySection
- createBendElement
- createBeamElement
- createLoadSet
- createNodeLoad
- createFreq1
- createHarmonicResponseAnalysis

- _create_set
- create_set_of_nodes
- create_set_of_elements
- create_set_of_materials
- create_set_of_properties
- create_set_of_outputs


## Delete
- deleteAllGeometry
- deleteAllMesh
- deleteAllOutput
- deleteAll
- deleteElement
- deleteNode
- delete_output


## Get
- getNodesInfoDialog
- getNodeInfo
- getElementsInfoDialog
- getElementInfo
- getPropertyInfo
- getNumberOfSet
- getSetValue
- getAnalysisStudies
- getEntityValue
- getFloatDialog
- get_app
- get_list_from_Femap_set

- _get_results
- get_element_results
- get_node_results

- get_dict_of_properties_from_element_set
- get_dict_of_output_titles_from_output_set
- get_dict_of_frequencies_from_output_set


## List
- listMaterial
- listProperties
- listSetIds


## Select
- selectNodesDialog
- selectElementsDialog
- selectProperty


## Other
- __init__
- printMessage
- message
- refineBeamElements
- replaceNodeInElement
- rebuild
- showMessageBox
- message
- import_output

"""

import pythoncom
import PyFemap
from PyFemap import constants
import sys
import numpy as np
import pandas as pd
import time



def display_progress_bar(iteration, total, bar_length=50, elapsed_time=None):
    progress = float(iteration) / total
    arrow = '=' * int(round(progress * bar_length) - 1) + '>'
    spaces = ' ' * (bar_length - len(arrow))

    if elapsed_time:
        print(f"\r[{arrow + spaces}] {int(progress * 100)}% Elapsed time: {elapsed_time}", end='')
    else:
        print(f"\r[{arrow + spaces}] {int(progress * 100)}%", end='')

    if progress == 1:
        print()

def convert_timer_to_readable(secs):
    mins, s = divmod(secs, 60)
    h, mins = divmod(mins, 60)

    h_str = f'{int(h)}h ' if h > 0 else ''
    mins_str = f'{int(mins)}m ' if mins > 0 else ''
    s_str = f'{round(s, 2)}s'

    return h_str + mins_str + s_str

def get_elapsed_time_readable(start_time):
    elapsed_time = time.time() - start_time

    return elapsed_time

class App:
    #Constructor Function, need to be called to set up the link with FEMAP and
    #display message "Python API Connected"
    def __init__(self):
        try:
            existObj = pythoncom.connect(PyFemap.model.CLSID)
            self._app = PyFemap.model(existObj)
            self._rbo = self._app.feResults
            self._app.feAppMessage(0,"Python API Connected")
        except:
            sys.exit("Femap doesn't appear to be opened")
    
    #Print Message
    def printMessage(self,message):
        self._app.feAppMessage(0,message)

    def message(self, color, message):
        '''
        0 = Normal
        1 = Highlight
        2 = Warning
        3 = Error
        '''
        self._app.feAppMessage(color, message)
    
    def deleteAllGeometry(self):
        self._app.feDeleteAll(True,False,False,True)
        
    def deleteAllMesh(self):
        self._app.feDeleteAll(False,True,False,True)
        
    def deleteAllOutput(self):
        self._app.feDeleteAll(False,False,True,True)
        
    def deleteAll(self):
        self._app.feDeleteAll(True,True,True,True)
    
    def deleteElement(self,id):
        elemSet = self._app.feSet
        elemSet.Add(id)
        self._app.feDelete(8, elemSet.ID)
        
    def deleteNode(self,id):
        nodeSet = self._app.feSet
        nodeSet.Add(id)
        self._app.feDelete(7, nodeSet.ID)
    
    # Function to create note. If no id is specified, Femap use the next empty ID
    def createNode(self,x,y,z,id=None):
        node = self._app.feNode
        if not id:
            id = node.NextEmptyID()
        node.x = x
        node.y = y
        node.z = z
        node.Put(id)
        
        return id
    
    # Function to create function. If no id is specified, Femap use the next empty ID
    def createFunction(self,fTitle,x_data,y_data,fType=None, id=None):
        function = self._app.feFunction
        if not id:
            id = function.NextEmptyID()
        if not fType:
            fType = 1
        function.title = fTitle
        function.type = fType
        function.PutFunctionList(len(x_data), x_data, y_data)
        function.Put(id)
        
        return id
    
    def selectNodesDialog(self,message="Select Nodes"):
        nodeSet = self._app.feSet
        rc = nodeSet.Select(7, True, message)
        nodeIds = []
        if rc ==-1 and nodeSet.Count()>0:
            rc = nodeSet.Reset()
            while True:
                nodeId = nodeSet.Next()
                if nodeId > 0:
                    nodeIds.append(nodeId)
                else:
                    break
        return nodeIds
    
    def getNodesInfoDialog(self,message="Select Nodes"):
        nodeIds = self.selectNodesDialog(message)
        nodes = []
        for nodeId in nodeIds:
            nodes.append(self.getNodeInfo(nodeId))
        return nodes
                
    def getNodeInfo(self,id):
        node = self._app.feNode
        node.Get(id)
        return [id,node.x,node.y,node.z]
    
    def selectElementsDialog(self, message="Select Elements"):
        elemSet = self._app.feSet
        rc = elemSet.Select(8, True, message)
        elemIds = []

        if rc == -1 and elemSet.Count() > 0:
            rc = elemSet.Reset()
            while True:
                elemId = elemSet.Next()
                if elemId > 0:
                    elemIds.append(elemId)
                else:
                    break

        return elemIds
    
    def getElementsInfoDialog(self,message="Select Elements"):
        elemIds = self.selectElementsDialog(message)
        elems = []
        for elemId in elemIds:
            elems.append(self.getElementInfo(elemId))
        return elems
                
    def getElementInfo(self,id):
        elem = self._app.feElem
        elem.Get(id)
        if elem.type == 5 or elem.type == 3: #Beam Element
            return {'ID':elem.ID,'type':elem.type,'propID':elem.propID,'topology':elem.topology,
                    'orientID':elem.orientID,'n1':list(elem.vnode)[0],'n2':list(elem.vnode)[1],'vorient':list(elem.vorient),'voffset':list(elem.voffset)}
        else:
            return {'ID':elem.ID,'type':elem.type,'propID':elem.propID,'topology':elem.topology,'vnode':list(elem.vnode)}
    
    def refineBeamElements(self,ids,ratio):
        elemSet = self._app.feSet
        for elemId in ids:
            elemSet.Add(elemId)
        print(ratio)
        self._app.feMeshRemesh(elemSet.ID,ratio,0,ratio,ratio,ratio,ratio,ratio,ratio,ratio,ratio,ratio,ratio,ratio,ratio)
    
    def selectProperty(self,message="Select Property"):
        nodeSet = self._app.feSet
        rc , id = nodeSet.SelectID(11, message)
        return id
    
    # Function to create standard material, Steel only defined for now, feel free to add
    # return the id of the material created
    def createStandardMaterial(self,name,type="Steel",id=None):
        mt = self._app.feMatl
        if not id:
            id = mt.NextEmptyID()
        mt.title = name
        if type == "Steel":
            mt.type = 0
            v = list(mt.mmat)
            v[0]=2.1e11
            v[6]=0.3
            v[49]=7850
            mt.mmat = v
        else:
            raise Exception("Unknown Material Type")
        mt.Put(id)
        return id
    
    def createIsotropicMaterial(self,name,modulus=None,poisson=None,density=None,damping=None,exp_coeff=None,refTemp=None,id=None):
        mt = self._app.feMatl
        if not id:
            id = mt.NextEmptyID()
        mt.title = name
        mt.type = 0
        v = list(mt.mmat)
        if modulus:
            v[0]=modulus
        if poisson:
            v[6]=poisson
        if density:
            v[49]=density
        if damping:
            v[50]=damping
        if exp_coeff:
            v[36]=exp_coeff
        if refTemp:
            v[51]=refTemp
        mt.mmat = v
        mt.Put(id)
        return id
    
    def createLumpedMassProperty(self,name,mass,id=None):
        pr = self._app.feProp
        pr.title = name
        pr.type = 27
        vpval = list(pr.vpval)
        vpval[7] = mass
        vpval[11] = mass
        vpval[12] = mass
        pr.vpval = vpval
        if not id:
            id = pr.NextEmptyID()
        pr.Put(id)
        return id
        
    def createLumpedMass(self,node,propertyId,id=None):
        feElem = self._app.feElem
        feElem.ClearNodeList(-1)
        feElem.type = 27
        feElem.propID = propertyId
        feElem.topology = 9
        vnode = list(feElem.vnode)
        vnode[0] = node
        feElem.vnode = vnode
        if not id:
            id = feElem.NextEmptyID()
        feElem.Put(id)
        return id
    
    def listMaterial(self):
        mats = []
        mtNb = self._app.feMatl.CountSet()
        
        if mtNb:
            mt = self._app.feMatl
            mt.Reset()
            for i in range(mtNb):
                mt.Next()
                mats.append({'id':mt.ID,'name':mt.title})
        return mats
    
    def listProperties(self):
        props = []
        propNb = self._app.feProp.CountSet()
        
        if propNb:
            prop = self._app.feProp
            prop.Reset()
            for i in range(propNb):
                prop.Next()
                props.append({'id':prop.ID,'name':prop.title})
        return props
    
    def getPropertyInfo(self,id=None):
        if not id:
            id = self.selectProperty()
        if id:
            pr = self._app.feProp
            pr.Get(id)
            return {'ID':pr.ID,'title':pr.title,'type':pr.type,'matlID':pr.matlID,'vflagI':list(pr.vflagI),'vpval':list(pr.vpval)}
        
    
    def createBeamPipeSection(self,name,matId,OD,thck,nsm=0,id=None):
        id = self.createBeamPropertySection(name,matId,6,OD/2,0,0,0,0,thck,nsm=0,id=None)
        return id
    
    def createBendB31Property(self,name,matId,OD,thck,radius,pressure=0,nsm=0,id=None):
        pr = self._app.feProp
        if not id:
            id = pr.NextEmptyID()
        pr.matlID = matId
        pr.type = 31 #bend
        vflagI = [5, 0, 0, 0]
        pr.vflagI = vflagI
        vpval = [float(OD), float(OD-2*thck), 0.0, 0.0, 0.0, 0.0, 0.0, float(nsm), 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, float(radius), 0.0, 0.0, float(pressure), 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        pr.vpval = vpval
        pr.title = name
        pr.Put(id)
        return id
        
    # Function to create a Beam property defined by it's section shape
    # Need to specify the material Id to be used and the standard shape type and dimension
    # Option to add non structural mass [Mass/Lenght]
    # return the id of the property created
    def createBeamPropertySection(self,name,matId,shape,dim1,dim2,dim3,dim4,dim5,dim6,nsm=0,id=None):
        pr = self._app.feProp
        if not id:
            id = pr.NextEmptyID()
        pr.matlID = matId
        pr.type = 5 #beam
        vflagI = list(pr.vflagI)
        vflagI[1] = shape
        pr.vflagI = vflagI
        vpval = list(pr.vpval)
        vpval[40] = dim1
        vpval[41] = dim2
        vpval[42] = dim3
        vpval[43] = dim4
        vpval[44] = dim5
        vpval[45] = dim6
        pr.vpval = vpval
        rc = pr.ComputeShape(False, False, False)
        vpval = list(pr.vpval)
        vpval[7] = nsm
        pr.vpval = vpval
        pr.title = name
        pr.Put(id)
        return id
    
    def createBendElement(self,propertyId,n1,n2,orientation,id=None):
        feElem = self._app.feElem
        if not id:
            id = feElem.NextEmptyID()
        feElem.ClearNodeList(-1)
        feElem.type = 31
        feElem.propID = propertyId
        feElem.topology = 0
        feElem.orientID = 0
        vnode = list(feElem.vnode)
        vnode[0] = n1
        vnode[1] = n2
        feElem.vnode = vnode
        feElem.vorient = [orientation[0],orientation[1],orientation[2]]
        feElem.Put(id)
        return id
    
    def replaceNodeInElement(self,elemId,nodePos,newNodeId): #/!\ 0-Index
        feElem = self._app.feElem
        feElem.Get(elemId)
        vnode = list(feElem.vnode)
        vnode[nodePos] = newNodeId
        feElem.vnode = vnode
        feElem.Put(elemId)
        
    

    def createBeamElement(self,propertyId,n1,n2,orientation,offset = None,id=None):
        """Function to create a Beam element
        Takes property ID, node 1 node 2 and orientation 
        Can optionally provide an offshet in format 

        Args:
            propertyId (int): _description_
            n1 (int): First node
            n2 (int): Second node
            orientation (tuple[x,y,z]): Element orientation vector
            offset (tuple[x1,y1,z1,x2,y2,z2], optional): Element offset for node one and node two. Defaults to None.
            id (int, optional): Element ID for the new element. Defaults to None.

        Returns:
            int: The created element ID.
        """
        
        feElem = self._app.feElem
        if not id:
            id = feElem.NextEmptyID()
        feElem.ClearNodeList(-1)
        feElem.type = 5
        feElem.propID = propertyId
        feElem.topology = 0
        feElem.orientID = 0
        vnode = list(feElem.vnode)
        vnode[0] = n1
        vnode[1] = n2
        feElem.vnode = vnode
        if offset:
            feElem.voffset = [offset[0],offset[1],offset[2],offset[3],offset[4],offset[5]]
        vorient = [float(orientation[0]),float(orientation[1]),float(orientation[2])]
        feElem.vorient = list(vorient)
        feElem.Put(id)
        return id
    
    def createLoadSet(self,title,id=None):
        feLoadSet = self._app.feLoadSet
        if not id:
            id = feLoadSet.NextEmptyID()
        feLoadSet.title=title
        feLoadSet.Put(id)
        return id
    
    def createNodeLoad(self,loadSetId,nodeId,x=None,y=None,z=None,function=None):
        feLoadSet = self._app.feLoadSet
        feLoadSet.Get(loadSetId)
        
        feLoadMesh = self._app.feLoadMesh
        feLoadDefinition = self._app.feLoadDefinition
        feLoadDefinition.loadType =constants.FLT_NFORCE 
        feLoadDefinition.dataType = constants.FT_SURF_LOAD
        feLoadDefinition.setID = feLoadSet.ID
        feLoadDefinition.Put(feLoadDefinition.NextEmptyID())
        
        feLoadMesh.meshID = nodeId
        feLoadMesh.CSys = 0
        
        feLoadMesh.type =constants.FLT_NFORCE
        feLoadMesh.setID = feLoadSet.ID
        
        if x:
            feLoadMesh.XOn = True
            feLoadMesh.x = x
        if y:
            feLoadMesh.YOn = True
            feLoadMesh.y = y
        if z:
            feLoadMesh.ZOn = True
            feLoadMesh.z = z
        if function:
            feLoadMesh.LoadFunction = function
        feLoadMesh.LoadDefinitionID = feLoadDefinition.ID
        feLoadMesh.Put(feLoadMesh.NextEmptyID())
        
        feLoadMesh.Active = feLoadMesh.ID
        
    def createFreq1(self,start,stop,step):
        id = self._app.feFreq.NextEmptyID()
        self._app.feFreq.AddFreq1(start,step,int((stop-start)/step),1,False)
        return id
        
    def createHarmonicResponseAnalysis(self,name,freqRange,load=0,constrain=0,overallDamping=0,displacement=-1,velocity=-1,stress=-1):
        analysis = self._app.feAnalysisMgr
        analysis.title = name
        analysis.Solver = 36
        analysis.AnalysisType = 4 
        analysis.NasDynDampOverall = overallDamping
        analysis.vBCSet = (constrain, 0, load, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        analysis.vOutput = (0,0,displacement,velocity,0,0,0,0,0,0,0,0,0,0,0,0,stress,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
        id = analysis.NextEmptyID()
        analysis.NasDynOn = True
        analysis.NasDynFreqID =id
        freqSet = self._app.feSet
        freqSet.Add(freqRange)
        analysis.PutUsedFREQS(freqSet.ID)
        analysis.Put(id)
        return id
    
    # Return of the number of sets in FEMAP Results
    def getNumberOfSet(self):
        return self._rbo.NumberOfSets()
    
    # Return the set value for the set id provided (usefull to get the frequency of a mode)
    def getSetValue(self,id):
        res = self._rbo.SetInfo(id)
        return res[3]
    
    def getAnalysisStudies(self):
        study = self._app.feAnalysisStudy
        self._rbo.NextStudyReset()
        studyId = self._rbo.NextStudy(0,1)
        cumulated = 1
        studies = []
        while studyId:
            name = self._rbo.StudyTitle(studyId)[1]
            study.Get(studyId)
            nbSet = study.CountOutputSets()
            studies.append((studyId,name,nbSet,cumulated))
            cumulated += nbSet
            studyId = self._rbo.NextStudy(0,1)
        return studies
    
    def listSetIds(self):
        setIds = []
        self._rbo.NextSetReset()
        setId = self._rbo.NextSet(0,1)
        while setId:
            setIds.append(setId)
            setId = self._rbo.NextSet(0,1)
        return setIds
    
    def getEntityValue(self,setID,vectorID,entityID):
        """Return the value of an element or node for a given vector ID at a given Set ID

        Args:
            setID (int): Output set ID
            vectorID (int): Output vector ID
            entityID (int): Node or element ID

        Returns:
            float: Value of the node or element for the given vector and output set IDs.
        """
        return self._rbo.EntityValueV2(setID,vectorID,entityID)[1]
    
    def rebuild(self):
        """Function to rebuild the model (usefull to update the display)"""
        self._app.feFileRebuild(False,True)
        
    def showMessageBox(self,message):
        self._app.feAppMessageBox(0,message)
        
    def getFloatDialog(self,message="Input number:",length=False,minVal=0,maxVal=1e20):
        if length:
            number = self._app.feGetRealLength(message)
        else:
            number = self._app.feGetReal(message,float(minVal),float(maxVal))
        if number[0]==-1:
            return number[1]
        else:
            return None
        
    def get_app(self):
        return self._app
    
    def message(self, color, message):
        
        """
        Set the colour using:
        0 = Normal
        1 = Highlight
        2 = Warning
        3 = Error
        """
        
        self._app.feAppMessage(color, message)
    
    def _create_set(self, entity_id, title, entity_list=None, select_all=False):
        entitySet = self._app.feSet
        if entity_list:
            _ = entitySet.AddArray(len(entity_list), entity_list)
        elif select_all:
            _ = entitySet.AddAll(entity_id)
        else:
            _ = entitySet.Select(entity_id, True, title)
        return entitySet
            
    def create_set_of_nodes(self, node_list=None, select_all=False):
        return self._create_set(7, 'Select Nodes', entity_list=node_list, select_all=select_all)
    
    def create_set_of_elements(self, element_list=None, select_all=False):
        return self._create_set(8, 'Select Elements', entity_list=element_list, select_all=select_all)
    
    def create_set_of_materials(self, material_list=None, select_all=False):
        return self._create_set(10, 'Select Materials', entity_list=material_list, select_all=select_all)
    
    def create_set_of_properties(self, property_list=None, select_all=False):
        return self._create_set(11, 'Select Properties', entity_list=property_list, select_all=select_all)
    
    def create_set_of_functions(self, property_list=None, select_all=False):
        return self._create_set(35, 'Select Functions', entity_list=property_list, select_all=select_all)
      
    def create_set_of_outputs(self, output_ids=None, select_all=False):
        outputSet = self._app.feSet
        if select_all:
            _ = outputSet.AddAll(28)
        elif output_ids:
            for set_id in output_ids:
                _ = outputSet.AddArray(len(output_ids), output_ids)
        else:
            _ = outputSet.SelectMultiIDV2(28, 1, 'Select Output Sets')
        return outputSet
    
    def feSetIn(self, feSet_raw):
        rc = feSet_raw.CopyToClipboard(True)
        attempts = 0
        while attempts < 3:
            try:
                from tkinter import Tk
                return list(map(int, Tk().selection_get(selection="CLIPBOARD").splitlines()))
            except Exception:
                attempts += 1
        self._app.feAppMessage(PyFemap.constants.FCM_NORMAL, "Failed to copy Set ID " + str(feSet_raw.ID) + "to Clipboard")
        return []
    
    def get_list_from_Femap_set(self, feSet):
        [_, _, entityList] = feSet.GetArray()
        return entityList
    


    def convert_nodal_stress(self, set_ids, elemental_vector, group, id=None):
        #feOutputSet_IDs = self.create_set_of_outputs(output_ids = set_ids)
        print("Converting elemental stress to nodal stress:")
        for idx, set_id in enumerate(set_ids, start=1):
            rc = self._app.feOutputProcessConvertV2(0, set_id, elemental_vector, 0, 0, group)
            # Display the progress bar
            display_progress_bar(idx, len(set_ids))
        print("Conversion complete!")
        return

    #This "get_results" extracts multiple output vectors across multiple locations and for multiple output sets, however the sets are output in the same column. To be merged with other one
    def _get_results_2(self, outputSet, vectors, entitySet, entityTypeID, **kwargs):
        """
        Parameters
        ----------
        outputSet : Femap Set
            Includes the output set ids.
        vectors : List
            Contains the output vector ids.
        entitySet : Femap Set
            Includes the entity ids.
        entityTypeID : Integer
            Femap API Entity Type (Reference Section 3.3.6 of Femap API 
            Documentation).
        **transform : String, optional
                - 'Nodal': transform nodal results to the node's output CSys

        Returns
        -------
        pandas dataframe
            Dataframe containing the results. Columns are:
                - 'id': Entity ID
                - 'set': output set
                - [vector id]: columns are the output vector ids
        """
        transform = kwargs.get('transform')
        results = []
        output = self._app.feResults
        [_, _, outputIDs] = outputSet.GetArray()
        entitySetID = entitySet.ID
        [_, nEntities, entityIDs] = entitySet.GetArray()
        nVectors = len(vectors)
        columnIndex = [i for i in range(nVectors)]
        print(f"entitySetID: {entitySetID}, nVectors: {nVectors}, columnIndex: {columnIndex}")

        for set_id in outputIDs:
            _ = output.clear()
            _ = output.DataNeeded(entityTypeID, entitySetID)
            
            for vect in vectors:
                [_, _, _] = output.AddColumnV2(set_id, vect, False)
            if transform == 'nodal':
                _ = output.SetNodalTransform(2, 0)
            _ = output.Populate()   
            [_, dVals, _] = output.GetRowsAndColumnsByID(entitySetID, nVectors, columnIndex)
            ret_values = output.GetRowsAndColumnsByID(entitySetID, nVectors, columnIndex)
            print(f"Return values from GetRowsAndColumnsByID: {ret_values}")

            print(f"Size of dVals: {len(dVals)}")
            print(f"nEntities: {nEntities}, nVectors: {nVectors}")

            dVals = np.reshape(dVals, [nEntities, nVectors])
            df = pd.concat([pd.DataFrame({'id': entityIDs, 'set': set_id}), 
                    pd.DataFrame(dVals, columns=vectors)], 
                    axis=1)     
            results.append(df)
        return pd.concat(results, ignore_index=True)

    #This "get_results" is good for extracting a single output vector across multiple locations and multiple output sets, i.e., stress over time or frequency at a weld. To be merged.
    def _get_results(self, outputSet, vectors, entitySet, entityTypeID, **kwargs):

        transform = kwargs.get('transform')
        results = []
        output = self._app.feResults
        [_, _, outputIDs] = outputSet.GetArray()
        entitySetID = entitySet.ID
        [_, nEntities, entityIDs] = entitySet.GetArray()

        # Assume only one vector for this revised method
        vect = vectors[0]

        data = {"id": entityIDs}

        #print("Fetching results from Femap:")

        for idx, set_id in enumerate(outputIDs, start=1):
            _ = output.clear()
            _ = output.DataNeeded(entityTypeID, entitySetID)
            [_, _, _] = output.AddColumnV2(set_id, vect, False)
            if transform == 'nodal':
                _ = output.SetNodalTransform(2, 0)
            _ = output.Populate()
            [_, dVals, _] = output.GetRowsAndColumnsByID(entitySetID, 1, [0])

            if dVals is None:
                raise ValueError(f"Failed to fetch results from Femap for output set {set_id}")

            # Add results for this output set to the data dictionary
            data[f"set_{set_id}"] = dVals

            # Display the progress bar
            display_progress_bar(idx, len(outputIDs))

        return pd.DataFrame(data)
    
    def get_element_results(self, outputSet, elementSet, vectors):
        '''
        outputSet = Femap set including output set ids
        elementSet = Femap set including element ids
        vectors = List of output vector ids
        '''
        return self._get_results(outputSet, vectors, elementSet, 8)
            
    def get_node_results(self, outputSet, nodeSet, vectors, transform=False):
        '''
        outputSet = Femap set including output set ids
        nodeSet = Femap set including node ids
        vectors = List of output vector ids
        transform = Boolian to transform results to output CSys
        '''
        tsfm = 'nodal' if transform else ''
        return self._get_results(outputSet, vectors, nodeSet, 7, transform=tsfm)
    
    def get_dict_of_properties_from_element_set(self, elementSet):
        element = self._app.feElem
        [_, _, entID, propID, _, _, _, _, _, _, _, _, _, _, _, _, _] = element.GetAllArray(elementSet.ID)
        propIDs = dict((e, p) for e, p in zip(entID, propID))
        return propIDs
    
    def get_dict_of_output_titles_from_output_set(self, outputSet):
        [_, _, outputIDs] = outputSet.GetArray()
        output = self._app.feResults
        titles = {}
        for set_id in outputIDs:
            [_, title] = output.SetTitle(set_id)
            titles[set_id] = title
        return titles
            
    def get_dict_of_frequencies_from_output_set(self, outputSet):
        [_, _, outputIDs] = outputSet.GetArray()
        output = self._app.feResults
        frequencies = {}
        for setID in outputIDs:
            [_, _, _, dSetValue] = output.SetInfo(setID)
            frequencies[setID] = dSetValue
        return frequencies
    
    def delete_output(self):
        self._app.feDeleteAll(False, False, True, True)
    
    def import_output(self, output_path):
        self._app.feFileReadNastranResults(0, output_path)
