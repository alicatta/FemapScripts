# -*- coding: utf-8 -*-
"""
WOOD VDN Femap API
"""

import pythoncom
"""
PyFemap is version specific, need to be generated once calling the following script:
import sys
from win32com.client import makepy
sys.argv = [“makepy”, “-o Pyfemap.py”, r“C:\ Program Files\Sie­mens\FEMAPvXXXXXXX\femap.tlb”] 
makepy.main() 
"""
import PyFemap
from PyFemap import constants
import sys


class App:
    """
    The main class to communicate with FEMAP
    """
    def __init__(self):
        """
        Constructor Function, need to be called to set up the link with FEMAP and
        display message "Python API Connected", stop the script if FEMAP is not opened
        """
        try:
            existObj = pythoncom.connect(PyFemap.model.CLSID)
            self._app = PyFemap.model(existObj)
            self._rbo = self._app.feResults
            self._app.feAppMessage(0,"Python API Connected")
        except:
            sys.exit("Femap doesn't appear to be opened")
            
    

    ##############################################################################
    #
    #        General FEMAP Functions
    #
    ##############################################################################

    
    def printMessage(self,message):
        '''
        Print Message in FEMAP Messages window.
        
        Parameters
        ----------
        message : string
            The message to display
        '''
        self._app.feAppMessage(0,message)   
        
    def showMessageBox(self,message):
        '''
        Display Message in FEMAP Message Popup Box.
        
        Parameters
        ----------
        message : string
            The message to display
        '''
        self._app.feAppMessageBox(0,message)
        
    def deleteAllGeometry(self):
        '''
        Delete All geometry objects in FEMAP.
        '''
        self._app.feDeleteAll(True,False,False,True)
        
    def deleteAllMesh(self):
        '''
        Delete All meshing objects in FEMAP.
        '''
        self._app.feDeleteAll(False,True,False,True)
        
    def deleteAllOutput(self):
        '''
        Delete All outputs (results) in FEMAP.
        '''
        self._app.feDeleteAll(False,False,True,True)
        
    def deleteAll(self):
        '''
        Delete everything in FEMAP (blank model).
        '''
        self._app.feDeleteAll(True,True,True,True)

    def rebuild(self):
        '''
        Rebuild the model (usefull to update the display)
        '''
        self._app.feFileRebuild(False,True)
        

    
    def getFloatDialog(self,message="Input number:",length=False,minVal=0,maxVal=1e20):
        '''
        Show Dialog window in FEMAP for user to input float number, if lenght option activated, user can use FEMAP measurement tool.
        
        Parameters
        ----------
        message : string, optional
            The message to display on dialog box title, default "input number"
        length : bool, optional
            If True, user can use the measuring tool, default False
        minVal : int or float, optional
            min value user can input, default 0
        maxVal : int or float, optional
            max value user can input, default 1e20
        '''
        if length:
            number = self._app.feGetRealLength(message)
        else:
            number = self._app.feGetReal(message,float(minVal),float(maxVal))
        if number[0]==-1:
            return number[1]
        else:
            return None
    ##############################################################################
    #
    #        Nodes Functions
    #
    ##############################################################################
    
    def createNode(self,x,y,z,id=None):
        '''
        Create a node at specified coordinate. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        x : int or float
            x coordinate
        y : int or float
            y coordinate
        z : int or float
            z coordinate
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of node created
        '''
        node = self._app.feNode
        if not id:
            id = node.NextEmptyID()
        node.x = x
        node.y = y
        node.z = z
        node.Put(id)
        
        return id
    
    def deleteNode(self,id):
        '''
        Delete a node for specified id.
        
        Parameters
        ----------
        id : int
        '''
        nodeSet = self._app.feSet
        nodeSet.Add(id)
        self._app.feDelete(7, nodeSet.ID)
        
    def selectNodesDialog(self,message="Select Nodes"):
        '''
        Show Dialog window in FEMAP for the user to select nodes, return the ids of the nodes selected
        
        Parameters
        ----------
        message : string, optional
            Text string to display in selection window
        Returns
        -------
        list of int
            list of ids for node selected
        '''
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
        '''
        Show Dialog window in FEMAP for the user to select nodes, return list of nodes infos tupple (id,x-coordinate,y-coordinate,z-coordinate)
        
        Parameters
        ----------
        message : string, optional
            Text string to display in selection window
        Returns
        -------
        list of tupple
            list of tupple containing node info (id,x-coordinate,y-coordinate,z-coordinate)
        '''
        nodeIds = self.selectNodesDialog(message)
        nodes = []
        for nodeId in nodeIds:
            nodes.append(self.getNodeInfo(nodeId))
        return nodes
                
    def getNodeInfo(self,id):
        '''
        Return nodes infos tupple (id,x-coordinate,y-coordinate,z-coordinate) for the provide node id
        
        Parameters
        ----------
        id : int
            node id
        Returns
        -------
        tupple
            containing node info (id,x-coordinate,y-coordinate,z-coordinate)
        '''
        node = self._app.feNode
        node.Get(id)
        return [id,node.x,node.y,node.z]
    
    ##############################################################################
    #
    #        Elements Functions
    #
    ##############################################################################          
    
    # Function to create a Beam element
    # Takes property ID, node 1 node 2 and orientation (format [x,y,z])
    # Can optionally provide an offshet in format [x1,y1,z1,x2,y2,z2]
    def createBeamElement(self,propertyId,n1,n2,orientation,offset = None,id=None):
        '''
        Create a a beam element based on the provided property id and 2 nodes ids. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        propertyId : int
            Property Id for the beam
        n1 : int 
            first beam node id
        n2 : int
            second beam node id
        orientation : list of float
            orientation vector coordinate [x,y,z]
        offset: list of float, optional
            offset vector coordinates for both nodes [x_n1,y_n1,z_n1,x_n2,y_n2,z_n2]
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of element created
        '''
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

    def createLumpedMass(self,propertyId,n1,id=None):
        '''
        Create a a Lumped Mass element based on the provided property id and a nodes id. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        propertyId : int
            Property Id for the lumped mass
        n1 : int 
            node id where to add the mass
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of element created
        '''
        feElem = self._app.feElem
        feElem.ClearNodeList(-1)
        feElem.type = 27
        feElem.propID = propertyId
        feElem.topology = 9
        vnode = list(feElem.vnode)
        vnode[0] = n1
        feElem.vnode = vnode
        if not id:
            id = feElem.NextEmptyID()
        feElem.Put(id)
        return id
    
    def createBendElement(self,propertyId,n1,n2,orientation,id=None):
        '''
        Create a a bend element based on the provided property id and 2 nodes ids. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        propertyId : int
            Property Id for the beam
        n1 : int 
            first beam node id
        n2 : int
            second beam node id
        orientation : list of float
            orientation vector coordinate [x,y,z]
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of element created
        '''
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
    
    def deleteElement(self,id):
        '''
        Delete an element for specified id.
        
        Parameters
        ----------
        id : int
        '''
        elemSet = self._app.feSet
        elemSet.Add(id)
        self._app.feDelete(8, elemSet.ID)
        
    
    def selectElementsDialog(self,message="Select Elements"):
        '''
        Show Dialog window in FEMAP for the user to select Elements, return the ids of the elements selected
        
        Parameters
        ----------
        message : string, optional
            Text string to display in selection window
        Returns
        -------
        list of int
            list of ids for elements selected
        '''
        elemSet = self._app.feSet
        rc = elemSet.Select(8, True, message)
        elemIds = []
        if rc ==-1 and elemSet.Count()>0:
            rc = elemSet.Reset()
            while True:
                elemId = elemSet.Next()
                if elemId > 0:
                    elemIds.append(elemId)
                else:
                    break
        return elemIds
    
    def getElementsInfoDialog(self,message="Select Elements"):
        '''
        Show Dialog window in FEMAP for the user to select nodes, return list of element infos dict (format depend on element type)
        
        Parameters
        ----------
        message : string, optional
            Text string to display in selection window
        Returns
        -------
        list of dict
            list of dict containing element info (format depend on element type)
        '''
        elemIds = self.selectElementsDialog(message)
        elems = []
        for elemId in elemIds:
            elems.append(self.getElementInfo(elemId))
        return elems
                
    def getElementInfo(self,id):
        '''
        Return element infos dict (format depend on element type) for the provided element id.
        
        Parameters
        ----------
        id : int
            element id
        Returns
        -------
        dict
            dict containing element info (format depend on element type)
        '''
        elem = self._app.feElem
        elem.Get(id)
        #TODO: Handle more element type
        if elem.type == 5 or elem.type == 3: #Beam Element
            return {'ID':elem.ID,'type':elem.type,'propID':elem.propID,'topology':elem.topology,
                    'orientID':elem.orientID,'n1':list(elem.vnode)[0],'n2':list(elem.vnode)[1],'vorient':list(elem.vorient),'voffset':list(elem.voffset)}
        else:
            return {'ID':elem.ID,'type':elem.type,'propID':elem.propID,'topology':elem.topology,'vnode':list(elem.vnode)}
    
    def replaceNodeInElement(self,elemId,nodePos,newNodeId): #/!\ 0-Index
        '''
        For a specific element, replace the node of the element by a new node id at the provide node position.
        
        Parameters
        ----------
        elemId : int
            Id of the element to modify
        nodePos : int
            position of the node to replace. /!\ zero-index. First node would have a 0 position.
        newNodeId : int
            Id of the new node
        '''
        feElem = self._app.feElem
        feElem.Get(elemId)
        vnode = list(feElem.vnode)
        vnode[nodePos] = newNodeId
        feElem.vnode = vnode
        feElem.Put(elemId)


    ##############################################################################
    #
    #        Material Functions
    #
    ##############################################################################    

    def createStandardMaterial(self,name,type="Steel",id=None):
        '''
        Create Material based on "Standard" property. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        name : string
            name to appear in femap for the material
        type : string,optional 
            type of standard material, default is steel
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of material created
        '''
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
        #TODO: Add more material
        else:
            raise Exception("Unknown Material Type")
        mt.Put(id)
        return id

    def createIsotropicMaterial(self,name,modulus=None,poisson=None,density=None,damping=None,exp_coeff=None,refTemp=None,id=None):
        '''
        Create Isotropic Material based on mechanical property provided. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        name : string
            name to appear in femap for the material
        modulus : float,optional 
            Young Modulus, if no input 0
        poisson : float,optional 
            Poisson Ratio, if no input 0
        density : float,optional 
            Material density, if no input 0
        damping : float,optional 
            Damping ratio, if no input 0
        exp_coeff : float,optional 
            Expension coefficient, if no input 0
        refTemp : float,optional 
            Ref temperature, if no input 0
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of material created
        '''
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
    
    def listMaterial(self):
        '''
        Return the list of material already defined in the FEMAP model as a list of dictionaries of id and name 
        

        Returns
        -------
        list of dictionaries
            return the list of material defined as disctionay containing the id and name of the material
        '''
        mats = []
        mtNb = self._app.feMatl.CountSet()
        
        if mtNb:
            mt = self._app.feMatl
            mt.Reset()
            for i in range(mtNb):
                mt.Next()
                mats.append({'id':mt.ID,'name':mt.title})
        return mats
    
    ##############################################################################
    #
    #        Properties Functions
    #
    ##############################################################################    
    
    def selectProperty(self,message="Select Property"):
        '''
        Show Dialog window in FEMAP for the user to select property, return the ids of the nodes selected
        
        Parameters
        ----------
        message : string, optional
            Text string to display in selection window
        Returns
        -------
        int
            id for property selected
        '''
        nodeSet = self._app.feSet
        rc , id = nodeSet.SelectID(11, message)
        return id
    

    
    def createLumpedMassProperty(self,name,mass,id=None):
        '''
        Create a a Lumped Mass property based on the provided mass value. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        name : int
            Name of the property
        mass : float 
            mass value
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of the property created
        '''
        #TODO: Manage inertia properties too
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
        

    
    def listProperties(self):
        '''
        Return the list of properties already defined in the FEMAP model as a list of dictionaries of id and name 
        

        Returns
        -------
        list of dictionaries
            return the list of properties defined as disctionay containing the id and name of the material
        '''
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
        '''
        Return property infos tupple (id,x-coordinate,y-coordinate,z-coordinate) for the provided node id (if no id is provided femap will prompt the user to select the id)
        
        Parameters
        ----------
        id : int,optional
            property id (if no id is provided femap will prompt the user to select the id)
        Returns
        -------
        tupple
            containing node info (ID,title,type,matID,vflagI,vpval). vflagI,vpval are specific properties to the type of properties
        '''
        #TODO: diferenciate between different type of common property so user don't have to deal with obscure vflagI,vpval variable
        if not id:
            id = self.selectProperty()
        if id:
            pr = self._app.feProp
            pr.Get(id)
            return {'ID':pr.ID,'title':pr.title,'type':pr.type,'matlID':pr.matlID,'vflagI':list(pr.vflagI),'vpval':list(pr.vpval)}

    # Function to create a Beam property defined by it's section shape
    # Need to specify the material Id to be used and the standard shape type and dimension
    # Option to add non structural mass [Mass/Lenght]
    # return the id of the property created
    def createBeamPropertySection(self,name,matId,shape,dim1,dim2,dim3,dim4,dim5,dim6,nsm=0,id=None):
        '''
        Create Beam property defined by it's section shape. Shape type and dimension follow NASTRAN PBEAM spec.
    
        
        Parameters
        ----------
        name : string
            name to appear in femap for the property
        matId : int
            material idea to be used
        shape : int 
            type of beam section (I,C channel...)
        dim1,dim2,dim3,dim4,dim5,dim6 : float
            dimension as per Nastran PBEAM spec
        nsm : float,optional 
            Non structural mass, if no input 0
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of material created
        '''
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
    
    def createBeamPipeSection(self,name,matId,OD,thck,nsm=0,id=None):
        '''
        Create Pipe property defined by it's OD and thickness.
    
        
        Parameters
        ----------
        name : string
            name to appear in femap for the property
        matId : int
            material idea to be used
        OD : float 
            outside diameter
        thck : float
            thickness
        nsm : float,optional 
            Non structural mass, if no input 0
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of material created
        '''
        id = self.createBeamPropertySection(name,matId,6,OD/2,0,0,0,0,thck,nsm=0,id=None)
        return id
    
    def createBendB31Property(self,name,matId,OD,thck,radius,pressure=0,nsm=0,id=None):
        '''
        Create Bend property defined by it's OD,thickness,radius and pressure.
    
        
        Parameters
        ----------
        name : string
            name to appear in femap for the property
        matId : int
            material idea to be used
        OD : float 
            outside diameter
        thck : float
            thickness
        radius : float
            bend radius
        pressure: float
            gauge pressure inside the bend
        nsm : float,optional 
            Non structural mass, if no input 0
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of material created
        '''
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
        

    ##############################################################################
    #
    #        Load Functions
    #
    ##############################################################################    


    def createLoadSet(self,title,id=None):
        '''
        Create Load Set.
    
        
        Parameters
        ----------
        title : string
            name to appear in femap for the Load Set
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of material created
        '''
        feLoadSet = self._app.feLoadSet
        if not id:
            id = feLoadSet.NextEmptyID()
        feLoadSet.title=title
        feLoadSet.Put(id)
        return id
    
    def createNodeLoad(self,loadSetId,nodeId,x=None,y=None,z=None,function=None):
        '''
        Create Load on a node within a given load set, load defined in 3 direction and eventuall a frequency function.
    
        
        Parameters
        ----------
        loadSetId : int
            id of the Load Set where the load will be
        nodeId : int
            id of the node where the load is defined
        x,y,z: float,optional
            load ampliude in x,y,z direction, if not specified 0 input.
        function: int,optional
            id of the frequency function associated to this load if applicable
        
        '''
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

    ##############################################################################
    #
    #        Function Functions
    #
    ##############################################################################  

    def createFunction(self,name,functionType,x,y,id=None):
        '''
        Create a function of the specified type. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        name : string
            function name
        functionType : int
            function type, 0=Dimensionless, 1=vs. Time, 2=vs. Temp, 3=vs. Freq, 4=vs. Stress...
        x : list of float or list of int
            x values
        y : list of float or list of int
                y values
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of node created
        '''
        if len(x) != len(y):
            raise Exception("createFunction: x and y list doesn't have the same lenght")
        function = self._app.feFunction
        if not id:
            id = function.NextEmptyID()
        function.title = name
        function.type = functionType
        function.PutFunctionList(len(x),x,y)
        function.Put(id)
        
        return id

    def createTimeFunction(self,name,x,y,id=None):
        '''
        Create a function of the specified type. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        name : string
            function name
        x : list of float or list of int
            x values
        y : list of float or list of int
                y values
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of node created
        '''
        return self.createFunction(name,1,x,y,id)
    
    def createFrequencyFunction(self,name,x,y,id=None):
        '''
        Create a function of the specified type. If no id is specified, Femap use the next empty ID
        
        Parameters
        ----------
        name : string
            function name
        x : list of float or list of int
            x values
        y : list of float or list of int
                y values
        id : int, optional
            if not specified, FEMAP use the next empty ID
        Returns
        -------
        int
            id of node created
        '''
        return self.createFunction(name,2,x,y,id)
    
    ##############################################################################
    #
    #        Analysis Functions
    #
    ##############################################################################  


    def createHarmonicResponseAnalysis(self,name,freqRange,load=0,constrain=0,overallDamping=0,displacement=-1,velocity=-1,stress=-1):
        '''
        Create a Harmonic Response analysis to next empty id
        
        Parameters
        ----------
        name : string
            analysis name
        freqRange : int
            Freq card id to be used for range of frequencies to perform the analysis
        load : int,optional
            Load case id
        constrain : int,optional
            Constrain case id
        overallDamping : float,optional
            Overall damping value if not defined in material
        displacement : int,optional
             displacement output group: -1 All model, 0 No output, 1,2,3... Group Id
        stress : int,optional
             stress output group: -1 All model, 0 No output, 1,2,3... Group Id
        velocity : int,optional
             velocity output group: -1 All model, 0 No output, 1,2,3... Group Id

        Returns
        -------
        int
            id of analysis created
        '''
        #TODO: More parameters for more flexibility
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
    
    def createFreq1(self,start,stop,step):
        '''
        Create a Frequency Range card (useful for frequency-domain analysis) defined by a a start frequency, stop frequency and a frequency step
        
        Parameters
        ----------
        start : float
            start Frequency
        stop : float
            stop Frequency
        step : float
            frequency step (e.g. 0.1 Hz)

        Returns
        -------
        int
            id of Freq Card created
        '''
        id = self._app.feFreq.NextEmptyID()
        self._app.feFreq.AddFreq1(start,step,int((stop-start)/step),1,False)
        return id
    
    
    
    ##############################################################################
    #
    #        Output/Result Functions
    #
    ############################################################################## 
    
    def getNumberOfSet(self):
        '''
        Get the total number of set available in result
        

        Returns
        -------
        int
            total number of set available in result
        '''
        return self._rbo.NumberOfSets()
    

    def getSetValue(self,id):
        '''
        Return the set value for the set id provided (usefull to get the frequency of a mode)
        
        Parameters
        ----------
        id : int
            set id
        
        Returns
        -------
        int,float
            value defined for given set (depend on the set type)
        '''
        res = self._rbo.SetInfo(id)
        return res[3]
    
    def getAnalysisStudies(self):
        '''
        Set can be grouped by studies. List all the studies available in femap model.
        
        Returns
        -------
        list of tupple
            each tupple contains the study id, the name, the nb of set it contains and the cumulated number of set since first study
        '''
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
        '''
        List all the set id avaiable in femap model (set might not be continiously numbered)
        
        Returns
        -------
        list of int
            list of all set id
        '''
        setIds = []
        self._rbo.NextSetReset()
        setId = self._rbo.NextSet(0,1)
        while setId:
            setIds.append(setId)
            setId = self._rbo.NextSet(0,1)
        return setIds
    
    # Return the Entity value (Element/Node for a given Vector ID at a given Set ID)
    def getEntityValue(self,setID,vectorID,entityID):
        '''
        Return a single Entity value (Element/Node for a given Vector ID at a given Set ID)
        
        Parameters
        ----------
        setID : int
            set id
        vectorID : int
            set id
        entityID : int
            entity id (either a element or node if, depending on the vector type)
        
        Returns
        -------
        float
            value associated to the selected vector for the selected node/element
        '''
        return self._rbo.EntityValueV2(setID,vectorID,entityID)[1]
    
