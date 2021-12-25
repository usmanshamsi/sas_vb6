Attribute VB_Name = "modInputProcedures"
'option base 1
Option Explicit

Sub subAddMaterial(Name As String, E As Double, G As Double, Alpha As Double)
On Error GoTo ErrorHandler

'-----------------------------------------------------------
'THE PROCEDURE ADDS ONE MATERIAL WITH SPECIFIED PROPERTIES
'-----------------------------------------------------------
    'CHECK IF A MATERIAL EXIST WITH NAME 'Name"
        If IsMaterialExist(Name, 0) = True Then
            MsgBox "Duplicate Material name passed to Add Material Procedure"
            ErrorOccured = True
            Exit Sub
        End If
        
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
    
    'INCREASE NUMBER OF MATERIALS BY 1
        NoOfMaterials = NoOfMaterials + 1
    
    'REDIMENSTION MATERIAL RELATED DATA ARRAYS
        ReDim Preserve MaterialName(NoOfMaterials)
        ReDim Preserve MatElasMod(NoOfMaterials)
        ReDim Preserve MatShearMod(NoOfMaterials)
        ReDim Preserve MatCoeffTher(NoOfMaterials)

    'ENTER THE MATERIAL DATA TO RELATED ARRAYS
        MaterialName(NoOfMaterials) = Name
        MatElasMod(NoOfMaterials) = E
        MatShearMod(NoOfMaterials) = G
        MatCoeffTher(NoOfMaterials) = Alpha
        
    'TRACK CHANGES
        ChangesSaved = False
        subPlotStr
        
Exit Sub
ErrorHandler:
    Call subDispErrInfo("adding material", Err.Number, Err.Description)


End Sub
Sub subEditMaterial(MaterialID As Long, NewName As String, NewE As Double, _
                    NewG As Double, NewAlpha As Double)
On Error GoTo ErrorHandler

'-----------------------------------------------------------
'THE PROCEDURE ADDS ONE MATERIAL WITH SPECIFIED PROPERTIES
'-----------------------------------------------------------
    'CHECK IF A MATERIAL EXIST WITH NAME 'NewName"
        If IsMaterialExist(NewName, MaterialID) = True Then
            MsgBox "Duplicate Material name passed to Edit Material Procedure"
            ErrorOccured = True
            Exit Sub
        End If
    'CHECK IF A MATERIAL ID IS VALID
        If MaterialID > NoOfMaterials Or MaterialID < 1 Then
            MsgBox "Invalid Material ID passed to Edit Material Procedure"
            ErrorOccured = True
            Exit Sub
        End If
    
    'ENTER NEW MATERIAL DATA TO RELATED ARRAYS
    
    Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        MaterialName(MaterialID) = NewName
        MatElasMod(MaterialID) = NewE
        MatShearMod(MaterialID) = NewG
        MatCoeffTher(MaterialID) = NewAlpha
        
    'TRACK CHANGES
        ChangesSaved = False
subPlotStr
Exit Sub
ErrorHandler:
    Call subDispErrInfo("editing material", Err.Number, Err.Description)

End Sub
Sub subAddSection(Name As String, Color As Long, Material As Long, _
                    Area As Double, Ix As Double, Iy As Double, j As Double)
On Error GoTo ErrorHandler
    
    'CHECK IF ANY SECTION EXIST WITH THE NAME 'Name'
        If IsSecExist(Name, 0) = True Then
            MsgBox "Duplicate Section name passed to Add Section Procedure"
            ErrorOccured = True
            Exit Sub
        End If
    
    'INCREASE NUMBER OF SECTIONS AND REDIM RELATED ARRAYS
    
    Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
    
        NoOfSections = NoOfSections + 1
        ReDim Preserve SecName(NoOfSections)
        ReDim Preserve SecColor(NoOfSections)
        ReDim Preserve SecMat(NoOfSections)
        ReDim Preserve SecArea(NoOfSections)
        ReDim Preserve SecIx(NoOfSections)
        ReDim Preserve SecIy(NoOfSections)
        ReDim Preserve SecJ(NoOfSections)
    
    'STORE NEW SECTION DATA
        SecName(NoOfSections) = Name
        SecColor(NoOfSections) = Color
        SecMat(NoOfSections) = Material
        SecArea(NoOfSections) = Area
        SecIx(NoOfSections) = Ix
        SecIy(NoOfSections) = Iy
        SecJ(NoOfSections) = j
        
    'TRACK CHANGES
        ChangesSaved = False
        subPlotStr
Exit Sub
ErrorHandler:
    Call subDispErrInfo("adding section", Err.Number, Err.Description)
End Sub
Sub subEditSection(SectionID As Long, NewName As String, NewColor As Long, NewMaterial As Long, _
                    NewArea As Double, NewIx As Double, NewIy As Double, NewJ As Double)

On Error GoTo ErrorHandler
    'CHECK IF ANY SECTION EXIST WITH THE NAME 'NewName'
        If IsSecExist(NewName, SectionID) = True Then
            MsgBox "Duplicate Section name passed to Edit Section Procedure"
            ErrorOccured = True
            Exit Sub
        End If
        
    'CHECK IF SECTION ID IS VALID
        If SectionID > NoOfSections Or SectionID < 1 Then
            MsgBox "Invalid SectionID passed to Add Section Procedure"
            ErrorOccured = True
            Exit Sub
        End If
      
    'MODIFY SECTION DATA
    
    Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
    
        SecName(SectionID) = NewName
        SecColor(SectionID) = NewColor
        SecMat(SectionID) = NewMaterial
        SecArea(SectionID) = NewArea
        SecIx(SectionID) = NewIx
        SecIy(SectionID) = NewIy
        SecJ(SectionID) = NewJ
        
    'TRACK CHANGES
        ChangesSaved = False
subPlotStr
Exit Sub
ErrorHandler:
    Call subDispErrInfo("editing section", Err.Number, Err.Description)
    
End Sub
Sub subAddNode(X As Double, Y As Double)

On Error GoTo ErrorHandler
'------------------------------------------------------------
'THE PROCEDURE ADDS ONE NODE AT (Xcoor,Ycoor)
'------------------------------------------------------------
    'CHECK WHETHER A NODE ALREADY EXIST AT THESE COORDINATES
    '--------------------------------------------------------
    If IsNodeExist(X, Y, 0, 0) = True Then
        MsgBox "A Node is already defined on these coordinates"
        ErrorOccured = True
        Exit Sub
    End If
    
    'INCREASE NUMBER OF NodeS BY 1
    '--------------------------------------------------------
    Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
    

    NoofNodes = NoofNodes + 1
    
    'REDIMENSION NODE RELATED DATA TO ACCOMODATE THE NEW ONE
    '--------------------------------------------------------
        'COORDINATE DATA
        '-----------------------
        ReDim Preserve Xcoor(NoofNodes)
        ReDim Preserve Ycoor(NoofNodes)
        'ReDim preserve Zcoor(NoofNodes)
        
        'TRANSLATIONAL RESTRAINTS
        '-----------------------
        ReDim Preserve TxRest(NoofNodes)
        ReDim Preserve TyRest(NoofNodes)
        'ReDim Preserve TzRest(NoOfNodes)
        
        'ROTATIONAL RESTRAINTS
        '-----------------------
        'ReDim Preserve RxRest(NoOfNodes)
        'ReDim Preserve RyRest(NoOfNodes)
        ReDim Preserve RzRest(NoofNodes)
        
        'TRANSLATIONAL FORCES
        '-----------------------
        ReDim Preserve XForce(NoofNodes)
        ReDim Preserve YForce(NoofNodes)
        'ReDim Preserve ZForce(NoOfNodes)
        
        'ROTATIONAL FORCES
        '-----------------------
        'ReDim Preserve XMom(NoOfNodes)
        'ReDim Preserve YMom(NoOfNodes)
        ReDim Preserve ZMom(NoofNodes)

    'ENTER NODAL DATA INTO COORDINATES ARRAYS
    '--------------------------------------------------------
        Xcoor(NoofNodes) = X
        Ycoor(NoofNodes) = Y
        'zcoor(noofnodes) = z
        
    'TRACK CHANGES
        ChangesSaved = False
        subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("adding node", Err.Number, Err.Description)
End Sub
Sub subEditNode(NodeNo As Long, NewX As Double, NewY As Double)

On Error GoTo ErrorHandler
'------------------------------------------------------------
'THE PROCEDURE REPLACES THE EXISTING COORDINATES OF 'NodoNo'
'WITH 'NewX' and 'NewY'
'------------------------------------------------------------
    'CHECK WHETHER ANOTHER NODE ALREADY EXIST AT THESE COORDINATES
    '--------------------------------------------------------
    If IsNodeExist(NewX, NewY, NodeNo, 0) = True Then
        MsgBox "Another Node is already defined on these coordinates"
        ErrorOccured = True
        Exit Sub
    End If
    
    'CHECK IF THE SPECIFIED NODE NO. EXISTS
    If NodeNo > NoofNodes Or NodeNo < 1 Then
        MsgBox "Invalid Node Number..."
        ErrorOccured = True
        Exit Sub
    End If
    
    'CHANGE THE COORDINATES OF NodeNo
    
    
    Xcoor(NodeNo) = NewX
    Ycoor(NodeNo) = NewY
    'Zcoor(nodeno)=NewZ
    
    'TRACK CHANGES
        ChangesSaved = False
        subPlotStr
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
Exit Sub
ErrorHandler:
Call subDispErrInfo("editing node", Err.Number, Err.Description)
End Sub

Sub subDeleteNode(NodeNo As Long)
On Error GoTo ErrorHandler

'------------------------------------------------------------
'THE PROCEDURE DELTES THE NODE NodeNo AND RENUMBERS REMAINING NODES
'------------------------------------------------------------
Dim i As Long
    'CHECK IF THE SPECIFIED NODE NO. EXISTS
    If NodeNo > NoofNodes Or NodeNo < 1 Then
        MsgBox "Invalid Node Number..."
        ErrorOccured = True
        Exit Sub
    End If
    
    'CHECK IF ANY ELEMENT IS CONNECTED TO NODE 'NodeNo'
    For i = 1 To NoOfElements
        If (NodeNo = Endi(i) Or NodeNo = Endj(i)) Then
            MsgBox "Element #" & i & " is connected to node#" & NodeNo & _
                    "." & vbCrLf & "Can not delete."
            Exit Sub
        End If
    Next i
    
    'EXCHANGE THE NODAL DATA
    For i = NodeNo To NoofNodes - 1
        Xcoor(i) = Xcoor(i + 1)
        Ycoor(i) = Ycoor(i + 1)
        'Zcoor(i)=zcoor(i+1)
        TxRest(i) = TxRest(i + 1)
        TyRest(i) = TyRest(i + 1)
        'TzRest(i) = TzRest(i + 1)
        'RxRest(i) = RxRest(i + 1)
        'RyRest(i) = RyRest(i + 1)
        RzRest(i) = RzRest(i + 1)
        XForce(i) = XForce(i + 1)
        YForce(i) = YForce(i + 1)
        'zforce(i) = zforce(i + 1)
        'XMom(i) = XMom(i + 1)
        'YMom(i) = ymom(i + 1)
        ZMom(i) = ZMom(i + 1)
        
    Next i
    
    'RENUMBER ELEMENT'S END NODES
    For i = 1 To NoOfElements
        
        If Endi(i) >= NodeNo Then Endi(i) = Endi(i) - 1
        If Endj(i) >= NodeNo Then Endj(i) = Endj(i) - 1
            
    Next i
    
    'DECREASE NUMBER OF NODES
    NoofNodes = NoofNodes - 1
    ReDim Preserve Xcoor(NoofNodes)
    ReDim Preserve Ycoor(NoofNodes)
    'ReDim Preserve zcoor(NoofNodes)
    ReDim Preserve TxRest(NoofNodes)
    ReDim Preserve TyRest(NoofNodes)
    'ReDim Preserve tzrest(NoofNodes)
    'ReDim Preserve RxRest(NoofNodes)
    'ReDim Preserve RyRest(NoofNodes)
    ReDim Preserve RzRest(NoofNodes)
    ReDim Preserve XForce(NoofNodes)
    ReDim Preserve YForce(NoofNodes)
    'ReDim Preserve zforce(NoofNodes)
    'ReDim Preserve xmom(NoofNodes)
    'ReDim Preserve ymom(NoofNodes)
    ReDim Preserve ZMom(NoofNodes)
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("deleting node", Err.Number, Err.Description)
End Sub
Sub subAddElement(Node1 As Long, Node2 As Long, section As Long)
On Error GoTo ErrorHandler

'------------------------------------------------------------
'THE PROCEDURE ADDS ONE ELMENT BETWEEN NODE1 AND NODE 2
'WITH SECTIONAL PROPERTIES DEFINED BY 'SECTION'
'------------------------------------------------------------

    'CHECK WHETHER NODE1, NODE2 AND SECTION EXISTS
        If Node1 > NoofNodes Or Node1 < 1 Then
            MsgBox "Initial Node do not exist..."
            ErrorOccured = True
            Exit Sub
        End If
    
        If Node2 > NoofNodes Or Node2 < 1 Then
            MsgBox "Final Node do not exist..."
            ErrorOccured = True
            Exit Sub
        End If
    
        If Node1 = Node2 Then
            MsgBox "Both node numbers can not be same"
            ErrorOccured = True
            Exit Sub
        End If
        
        If section > NoOfSections Or section < 1 Then
            MsgBox "Invalid Section passed to Add Element Procedure", , "Invalid Data"
            ErrorOccured = True
            Exit Sub
        End If
    
    'CHECK IF THE ELEMENT IS ALREADY DEFINED BETWEEN NODE1 AND NODE2
        If IsMemExist(Node1, Node2, 0) = True Then
                MsgBox "An Element are already present between nodes " & Node1 & " and " & Node2
                ErrorOccured = True
                Exit Sub
        End If
    
    'INCREASE NUMBER OF ELEMENTS BY 1
        NoOfElements = NoOfElements + 1
    
    'REDIMENSION ELEMENT RELATED DATA ARRAYS
        ReDim Preserve Endi(NoOfElements)
        ReDim Preserve Endj(NoOfElements)
        ReDim Preserve AsgnSec(NoOfElements)
        ReDim Preserve ElemLoadi(NoOfElements)
        ReDim Preserve ElemLoadj(NoOfElements)
        ReDim Preserve ElemALoadi(NoOfElements)
        ReDim Preserve ElemALoadj(NoOfElements)
        
    'FINALLY ADD THE ELEMENT
        Endi(NoOfElements) = Node1
        Endj(NoOfElements) = Node2
        AsgnSec(NoOfElements) = section
        
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("adding element", Err.Number, Err.Description)
End Sub
Sub subEditElement(ElemNo As Long, NewNode1 As Long, NewNode2 As Long, NewSection As Long)
On Error GoTo ErrorHandler

'------------------------------------------------------------
'THE PROCEDURE ASSIGNS NEW VALUES OF NODES AND SECTION TO
'ELEMENT NO 'ElemNo'
'------------------------------------------------------------

    'CHECK WHETHER NODE1, NODE2 AND SECTION EXISTS
        If NewNode1 > NoofNodes Or NewNode1 < 1 Then
            MsgBox "Initial Node do not exist..."
            ErrorOccured = True
            Exit Sub
        End If
    
        If NewNode2 > NoofNodes Or NewNode2 < 1 Then
            MsgBox "Final Node do not exist..."
            ErrorOccured = True
            Exit Sub
        End If
    
        If NewNode1 = NewNode2 Then
            MsgBox "Both node numbers can not be same"
            ErrorOccured = True
            Exit Sub
        End If
        
        If NewSection > NoOfSections Or NewSection < 1 Then
            MsgBox "Invalid Section passed to Add Element Procedure", , "Invalid Data"
            ErrorOccured = True
            Exit Sub
        End If
    
    'CHECK IF THE ELEMENT IS ALREADY DEFINED BETWEEN NODE1 AND NODE2
        If IsMemExist(NewNode1, NewNode2, ElemNo) = True Then
                MsgBox "An Element are already present between nodes " & NewNode1 & " and " & NewNode2
                ErrorOccured = True
                Exit Sub
        End If
    
    'CHECK IF THE ELEMENT 'ElemNo' EXIST
        If ElemNo > NoOfElements Or ElemNo < 1 Then
            MsgBox "Element do not exist..."
            ErrorOccured = True
            Exit Sub
        End If
        
    'FINALLY MODIFY THE ELEMENT
        Endi(ElemNo) = NewNode1
        Endj(ElemNo) = NewNode2
        AsgnSec(ElemNo) = NewSection
        
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("editing element", Err.Number, Err.Description)
End Sub


Sub subDeleteElement(ElemNo As Long) ', NewNode1 As Long, NewNode2 As Long, NewSection As Long)
On Error GoTo ErrorHandler

'------------------------------------------------------------
'THE PROCEDURE ASSIGNS NEW VALUES OF NODES AND SECTION TO
'ELEMENT NO 'ElemNo'
'------------------------------------------------------------
Dim i As Long
    'CHECK IF THE ELEMENT 'ElemNo' EXIST
        If ElemNo > NoOfElements Or ElemNo < 1 Then
            MsgBox "Element do not exist..."
            ErrorOccured = True
            Exit Sub
        End If
        
    'EXCHANGE THE ELEMENT PROPERTIES
    For i = ElemNo To NoOfElements - 1
        Endi(i) = Endi(i + 1)
        Endj(i) = Endj(i + 1)
        AsgnSec(i) = AsgnSec(i + 1)
        ElemLoadi(i) = ElemLoadi(i + 1)
        ElemLoadj(i) = ElemLoadj(i + 1)
        ElemALoadi(i) = ElemALoadi(i + 1)
        ElemALoadj(i) = ElemALoadj(i + 1)
    Next i
        
    'DECREASE NO OF ELEMENTS
    NoOfElements = NoOfElements - 1
    ReDim Preserve Endi(NoOfElements)
    ReDim Preserve Endi(NoOfElements)
    ReDim Preserve Endi(NoOfElements)
    ReDim Preserve ElemLoadi(NoOfElements)
    ReDim Preserve ElemLoadj(NoOfElements)
    ReDim Preserve ElemALoadi(NoOfElements)
    ReDim Preserve ElemALoadj(NoOfElements)

    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("deleting element", Err.Number, Err.Description)
End Sub

Sub subAddNodalLoad(NodeNo As Long, Fx As Double, Fy As Double, _
                                    Mz As Double)
On Error GoTo ErrorHandler
'--------------------------------------------------------------
'THE PROCEDURE ADDS THE NODAL FORCES INTO EXISTING NODAL FORCES
'AT NODE 'NodeNo'
'--------------------------------------------------------------

    'CHECK IF THE NODE NUMBER IS VALID
        If NodeNo > NoofNodes Or NodeNo < 1 Then
                MsgBox "Invalid node number passed to Add Nodal Load Procedure."
                ErrorOccured = True
                Exit Sub
        End If
    
    'ADD NODAL FORCES
        XForce(NodeNo) = XForce(NodeNo) + Fx
        YForce(NodeNo) = YForce(NodeNo) + Fy
        'ZForce(NODENO) = ZForce(NODENO) + FZ
        'XMom(NODENO) = XMom(NODENO) + MX
        'YMom(NODENO) = YMom(NODENO) + MY
        ZMom(NodeNo) = ZMom(NodeNo) + Mz
        
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("adding nodal load", Err.Number, Err.Description)
End Sub

Sub subReplaceNodalLoad(NodeNo As Long, Fx As Double, Fy As Double, _
                                    Mz As Double)
On Error GoTo ErrorHandler
'--------------------------------------------------------------
'THE PROCEDURE REPLACES THE NODAL FORCES AT NODE 'NodeNo'
'--------------------------------------------------------------

    'CHECK IF THE NODE NUMBER IS VALID
        If NodeNo > NoofNodes Or NodeNo < 1 Then
                MsgBox "Invalid node number passed to Replace Nodal Load Procedure."
                ErrorOccured = True
                Exit Sub
        End If
    
    'ADD NODAL FORCES
        XForce(NodeNo) = Fx
        YForce(NodeNo) = Fy
        'ZForce(NODENO) = FZ
        'XMom(NODENO) = MX
        'YMom(NODENO) = MY
        ZMom(NodeNo) = Mz
        
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("replacing nodal load", Err.Number, Err.Description)
End Sub
Sub subApplyRestraint(NodeNo As Long, Tx As Integer, Ty As Integer, Rz As Integer)
On Error GoTo ErrorHandler

'TO ADD SUPPORT CONDITIONS TO NODES
'----------------------------------

    'CHECK IF AN EXISTING NODE IS SPECIFIED
        If NodeNo > NoofNodes Or NodeNo < 1 Then
            MsgBox "Invalid node passed to Apply Restraint Procedure"
            ErrorOccured = True
            Exit Sub
        End If
    
    'APPLY CONDITIONS
        TxRest(NodeNo) = Tx
        TyRest(NodeNo) = Ty
        'TzRest(Nodeno) = TZ
        'RxRest(Nodeno) = RX
        'RyRest(Nodeno) = RY
        RzRest(NodeNo) = Rz
        
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
Exit Sub
ErrorHandler:
Call subDispErrInfo("applying restrains", Err.Number, Err.Description)
End Sub

Sub subAddElemLoad(ElemNo As Long, UDL1 As Double, UDL2 As Double, UDAL1 As Double, UDAL2 As Double)
On Error GoTo ErrorHandler

'--------------------------------------------------------------
'THE PROCEDURE ADDS THE ELEMENT FORCES INTO EXISTING ELEMENT FORCES
'AT ELEMENT 'ElemNo'
'--------------------------------------------------------------

    'CHECK IF THE ELEMENT NUMBER IS VALID
        If ElemNo > NoOfElements Or ElemNo < 1 Then
                MsgBox "Invalid element number passed to Add Element Load Procedure."
                ErrorOccured = True
                Exit Sub
        End If
    
    'ADD Element FORCES
        ElemLoadi(ElemNo) = ElemLoadi(ElemNo) + UDL1
        ElemLoadj(ElemNo) = ElemLoadj(ElemNo) + UDL2
        ElemALoadi(ElemNo) = ElemALoadi(ElemNo) + UDAL1
        ElemALoadj(ElemNo) = ElemALoadj(ElemNo) + UDAL2
        
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
        
Exit Sub
ErrorHandler:
Call subDispErrInfo("adding element load", Err.Number, Err.Description)
End Sub

Sub subReplaceElemLoad(ElemNo As Long, UDL1 As Double, UDL2 As Double, UDAL1 As Double, UDAL2 As Double)
On Error GoTo ErrorHandler

'--------------------------------------------------------------
'THE PROCEDURE REPLACES THE EXISTING ELEMENT FORCES
'--------------------------------------------------------------

    'CHECK IF THE ELEMENT NUMBER IS VALID
        If ElemNo > NoOfElements Or ElemNo < 1 Then
                MsgBox "Invalid element number passed to Add Element Load Procedure."
                ErrorOccured = True
                Exit Sub
        End If
    
    'ADD Element FORCES
        ElemLoadi(ElemNo) = UDL1
        ElemLoadj(ElemNo) = UDL2
        ElemALoadi(ElemNo) = UDAL1
        ElemALoadj(ElemNo) = UDAL2
        
    'TRACK CHANGES
        ChangesSaved = False
        Analyzed = False: PlotSF = False: PlotBMD = False: PlotAxial = False: PlotDeflectedShape = False
        subPlotStr
        
Exit Sub
ErrorHandler:
Call subDispErrInfo("replacing/deleting element load", Err.Number, Err.Description)
End Sub

