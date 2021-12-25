Attribute VB_Name = "modFileOpenSave"
'option base 1
Option Explicit

Sub subOpenFile(filename As String)
On Error GoTo ErrorHandler
'-----------
Dim i As Long, temp As String
    Analyzed = False
    'INITIALIZE FILE
    '--------------------------------------
    Open filename For Input As #1
    
        'READ FILE NAME
        '--------------------------------------
        Input #1, filename
        Line Input #1, temp
        
        'READ NODAL DATA
        '--------------------------------------
        Input #1, temp

        Input #1, NoofNodes
        Line Input #1, temp
        ReDim Xcoor(NoofNodes)
        ReDim Ycoor(NoofNodes)

        ReDim TxRest(NoofNodes)
        ReDim TyRest(NoofNodes)
        ReDim RzRest(NoofNodes)
        ReDim XForce(NoofNodes)
        ReDim YForce(NoofNodes)
        ReDim ZMom(NoofNodes)
        
        For i = 1 To NoofNodes
            Input #1, Xcoor(i), Ycoor(i), TxRest(i), TyRest(i), RzRest(i), XForce(i), YForce(i), ZMom(i)
        Next i
        
        Input #1, temp
                       
                
        'READ ELEMENT DATA
        '--------------------------------------
        Input #1, temp
        Input #1, NoOfElements
        Line Input #1, temp

        ReDim Endi(NoOfElements)
        ReDim Endj(NoOfElements)
        ReDim AsgnSec(NoOfElements)
        ReDim ElemLoadi(NoOfElements)
        ReDim ElemLoadj(NoOfElements)
        ReDim ElemALoadi(NoOfElements)
        ReDim ElemALoadj(NoOfElements)
        
        For i = 1 To NoOfElements
            Input #1, Endi(i), Endj(i), AsgnSec(i), ElemLoadi(i), ElemLoadj(i), ElemALoadi(i), ElemALoadj(i)
        Next i
        Input #1, temp
        
        'READ MATERIAL DATA
        '--------------------------------------
        Input #1, temp
        Input #1, NoOfMaterials
        Line Input #1, temp
        
        ReDim MaterialName(NoOfMaterials)
        ReDim MatElasMod(NoOfMaterials)
        ReDim MatShearMod(NoOfMaterials)
        ReDim MatCoeffTher(NoOfMaterials)
        
        For i = 1 To NoOfMaterials
            Input #1, MaterialName(i), MatElasMod(i), MatShearMod(i), MatCoeffTher(i)
        Next i
        Input #1, temp
    
        'READ SECTION DATA
        '--------------------------------------
        Input #1, temp
        Input #1, NoOfSections
        Line Input #1, temp
        
        ReDim SecName(NoOfSections)
        ReDim SecColor(NoOfSections)
        ReDim SecMat(NoOfSections)
        ReDim SecIx(NoOfSections)
        ReDim SecIy(NoOfSections)
        ReDim SecArea(NoOfSections)
        ReDim SecJ(NoOfSections)
        
        For i = 1 To NoOfSections
            Input #1, SecName(i), SecMat(i), SecIx(i), SecIy(i), SecArea(i), SecJ(i), SecColor(i)
        Next i
        Input #1, temp
        
    'CLOSE FILE
    '-------------------------------------------
    Close #1
'------------
   
    ChangesSaved = True
    FileSaved = True
    

Exit Sub
ErrorHandler:
MsgBox Err.Number
MsgBox Err.Description
    Call subDispErrInfo("opening file", Err.Number, Err.Description)
Close #1
End Sub



Sub subSaveFile(filename As String)
On Error GoTo ErrorHandler
'----------------------------
Dim i As Long
    
    'INITIALIZE FILE
    '--------------------------------------
    Open filename For Output As #1
    
        'SAVE FILE NAME
        '--------------------------------------
        Write #1, filename
        Print #1,
        
        'SAVE NODAL DATA
        '--------------------------------------
        Write #1, "NODAL DATA"
        Write #1, NoofNodes
        Write #1, "X", "Y", "TxRest", "TyRest", "RzRest", "XForce", "YForce", "ZMom"
        For i = 1 To NoofNodes
            Write #1, Xcoor(i), Ycoor(i), TxRest(i), TyRest(i), RzRest(i), XForce(i), YForce(i), ZMom(i)
        Next i
        Print #1,
        
        'SAVE ELEMENT DATA
        '--------------------------------------
        Write #1, "ELEMENT DATA"
        Write #1, NoOfElements
        Write #1, "Node1", "Node2", "Section", "UDL1", "UDL2", "UDAL1", "UDAL2"
        For i = 1 To NoOfElements
            Write #1, Endi(i), Endj(i), AsgnSec(i), ElemLoadi(i), ElemLoadj(i), ElemALoadi(i), ElemALoadj(i)
        Next i
        Print #1,
        
        'SAVE MATERIAL DATA
        '--------------------------------------
        Write #1, "MATERIALS"
        Write #1, NoOfMaterials
        Write #1, "Name", "E", "G", "Alpha"
        For i = 1 To NoOfMaterials
            Write #1, MaterialName(i), MatElasMod(i), MatShearMod(i), MatCoeffTher(i)
        Next i
        Print #1,
    
        'SAVE SECTION DATA
        '--------------------------------------
        Write #1, "Section Data"
        Write #1, NoOfSections
        Write #1, "Name", "Material", "I-x", "I-y", "Area", " J ", "Color"
        For i = 1 To NoOfSections
            Write #1, SecName(i), SecMat(i), SecIx(i), SecIy(i), SecArea(i), SecJ(i), SecColor(i)
        Next i
        Print #1,
        
    'CLOSE FILE
    '-------------------------------------------
    Close #1
    
    FileSaved = True
    ChangesSaved = True
    
Exit Sub
ErrorHandler:

Call subDispErrInfo("saving file", Err.Number, Err.Description)
Close #1

End Sub

