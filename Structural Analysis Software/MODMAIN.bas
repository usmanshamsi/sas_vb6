Attribute VB_Name = "modMain"
'option base 1
Option Explicit

'GLOBAL VARIABLE DECLARATION
'----------------------------
    Public ChangesSaved As Boolean      'TRACK WHETHER THE CHANGES MADE ARE SAVED
    Public ErrorOccured As Boolean      'TRACK WHETHER AN ERROR HAS BEEN ENCOUNTERED
                                        '...AND NOT DEALT WITH
    Public FileSaved As Boolean         'CHECK WHETHER A FILE NAME IS ASSIGNED TO CURRENT PROJECT
    Public filename As String           'Name of current file
    Public FileTitle As String
    Public Analyzed As Boolean
    Public PDOFs As Long, DOFs As Long
    Public PlotDeflectedShape As Boolean, DeflectionScale As Double
    Public PlotBMD As Boolean, BMDScale As Double
    Public PlotSF As Boolean, SFScale As Double
    Public plotLoads As Boolean, LoadScale As Double
    Public PlotAxial As Boolean, AxialScale As Double
    Public Mesh As Integer

'VARIABLES FOR NODAL DATA
'--------------------------
    Public NoofNodes As Long                        '# OF NODES
    ---Public Xcoor() As Double, Ycoor() As Double, Zcoor() As Double       'COORDINATE DATA
    ---Public DXcoor() As Double, DYcoor() As Double, DZcoor() As Double
    ???Public INIDIS1() As Double, INIDIS2() As Double
    ---Public TxRest() As Integer, TyRest() As Integer, TzRest() As Integer 'TRANSLATIONAL RESTRAINTS
    ---Public RxRest() As Integer, RyRest() As Integer, RzRest() As Integer 'ROTATIONAL RESTRAINTS
    ---Public XForce() As Double, YForce() As Double, ZForce() As Double   'TRANSLATIONAL FORCES
    ---Public XMom() As Double, YMom() As Double, ZMom() As Double        'ROTATIONAL FORCES / MOMENTS

'VARIABLES FOR ELEMENT DATA
'---------------------------
    Public NoOfElements As Long         '# OF ELEMENTS
    Public Endi() As Long, Endj() As Long       'END NODES OF ELEMENTS
    Public AsgnSec() As Long            'SECTION ID ASSIGNED TO ELEMENT
    Public ElemLoadi() As Double, ElemLoadj() As Double, ElemALoadi() As Double, ElemALoadj() As Double
                        'VALUE OF DISTRIBUTED LOAD ON NODE1 AND NODE2
    
    Public ShearForce() As Double, SFX() As Double, SFY() As Double
    Public BendingMoment() As Double, BMX() As Double, BMY() As Double
    Public AxialForce() As Double, AxialX() As Double, AxialY() As Double
    Public Slope() As Double
    Public DeflectionX() As Double, DeflectionY() As Double, Deflection() As Double
    
    
'VARIABLES FOR MATERIAL DATA
'----------------------------
    Public NoOfMaterials As Long        '# OF MATERIALS
    Public MaterialName() As String     'NAME OF MATERIAL
    Public MatElasMod() As Double, MatShearMod() As Double, MatCoeffTher() As Double
                        'E, G AND THERMAL ELONGATION COEFFICIENTS
    
'VARIABLES FOR SECTION DATA
'---------------------------
    Public NoOfSections As Long         '# OF SECTIONS
    Public SecName() As String          'NAME OF SECTION
    Public SecColor() As Long
    Public SecMat() As Long
    Public SecIx() As Double, SecIy() As Double, SecArea() As Double, SecJ() As Double
                        'AREA AND MOMENT OF INERTIAS OF SECTION ABOUT
                        'X, Y AND POLAR AXES
                        
'VARIABLES TO STORE FORCE, DISPLACEMENTS, STIFFNESS AND
'...TRANSFORMATION MATRICES
'---------------------------------------------------------
    Public ElemLocalStiff() As Double      'LOCAL K OF ELEMENTS
    Public ElemTrans() As Double           'TRANSFORMATION MATRIX OF ELEMENTS
    Public ElemGlobalStiff() As Double     'GLOBAL K OF ELEMENTS T'*K(LOCAL)*T
    Public SysStiff() As Double            'GLOBAL K FOR WHOLE STRUCTURE
    Public DispVector() As Double          'DISPLACEMENT VECTOR
    Public ForceVector() As Double         'FORCE VECTOR
    Public EndActions() As Double
    Public DirectForce() As Double
    Public RestVector() As Double          'Restraints vector
    Public RedSysStiff() As Double         'REDUCED K FOR WHOLE SYSTEM AFTER
                                           '...SUPPORT CONDITIONS
                                           
Sub Main()
On Error GoTo ErrorHandler

'THE SOFTWARE STARTS WITH THIS SUBROUTINE
''---------------------------------------
    'LOAD DEFAULT VALUES IN MEMORY
        LoadDefVal
    'DISPLAY THE MAIN WINDOW
        frmMain.Show
Exit Sub
ErrorHandler:
Call subDispErrInfo("in the main module", Err.Number, Err.Description)
End Sub

Public Sub LoadDefVal()

On Error GoTo ErrorHandler

ErrorOccured = False
ChangesSaved = True
FileSaved = False
filename = ""
FileTitle = "Untitled"
Analyzed = False
NoOfElements = 0
NoofNodes = 0
NoOfMaterials = 0
NoOfSections = 0
Mesh = 50

ReDim TxRest(0): ReDim TyRest(0): ReDim RzRest(0)
ReDim XForce(0): ReDim YForce(0): ReDim ZMom(0)
ReDim ElemLoadi(0): ReDim ElemLoadj(0): ReDim ElemALoadi(0): ReDim ElemALoadj(0)


'ADD DEFAULT MATERIALS
'------------------------
    Call subAddMaterial("Concrete 4ksi", 3600, 1500, 0.000001)
    Call subAddMaterial("Concrete 3ksi", 3122, 1300, 0.000001)
    Call subAddMaterial("Steel 36", 29000, 11150, 0.0001)
    
'For Sections

    Call subAddSection("Conc 1x1", vbRed, 1, 1, 1, 1, 1)


ChangesSaved = True

Exit Sub
ErrorHandler:
Call subDispErrInfo("loading default values", Err.Number, Err.Description)
End Sub



Function IsNodeExist(X As Double, Y As Double, Exception As Long, Tolerance) As Boolean
On Error GoTo ErrorHandler

'-----------------------------------------------------------------
'THE FUNCTION CHECKS WHETHER A NODE IS ALREADY DEFINED ON THE
'THE COORDINATES X,Y WITH IN DISTANCE DEFINED BY TOLERANCE
'-----------------------------------------------------------------
Dim X1, Y1 As Double
Dim i As Long
    For i = NoofNodes To 1 Step -1
        X1 = Xcoor(i)
        Y1 = Ycoor(i)
        If (X = X1 And Y = Y1 And (Not i = Exception)) Then
            IsNodeExist = True
            Exit Function
        End If
    Next i
IsNodeExist = False

Exit Function
ErrorHandler:
Call subDispErrInfo("IsNodeExist procedure", Err.Number, Err.Description)
End Function

Function IsMemExist(Node1 As Long, Node2 As Long, Exception As Long) As Boolean
On Error GoTo ErrorHandler

'-----------------------------------------------------------------
'THE FUNCTION CHECKS WHETHER AN ELEMENT ALREADY EXISTS BETWEEN
'NODE1 AND NODE2
'-----------------------------------------------------------------
Dim X, Y, i As Long
    For i = NoOfElements To 1 Step -1
        X = Endi(i)
        Y = Endj(i)
        If (Node1 = X And Node2 = Y And i <> Exception) Or _
           (Node2 = X And Node1 = Y And i <> Exception) Then
            IsMemExist = True
            Exit Function
        End If
    Next i
IsMemExist = False

Exit Function
ErrorHandler:
Call subDispErrInfo("IsMemberExist procedure", Err.Number, Err.Description)
End Function

Function IsMaterialExist(Name As String, Exception As Long) As Boolean

On Error GoTo ErrorHandler
'-----------------------------------------------------------------
'THE FUNCTION CHECKS WHETHER A MATERIAL IS ALREADY DEFINED WITH
'A NAME 'Name"
'-----------------------------------------------------------------
Dim i As Long
Dim X As String
    For i = 1 To NoOfMaterials
        X = MaterialName(i)
            If (X = Name And i <> Exception) Then
                IsMaterialExist = True
                Exit Function
            End If
    Next i
IsMaterialExist = False

Exit Function
ErrorHandler:
Call subDispErrInfo("IsMaterialExist procedure", Err.Number, Err.Description)
End Function


Function IsSecExist(Name As String, Exception As Long) As Boolean

On Error GoTo ErrorHandler
'-----------------------------------------------------------------
'THE FUNCTION CHECKS WHETHER A SECTION IS ALREADY DEFINED WITH
'A NAME 'Name"
'-----------------------------------------------------------------
Dim X As String
Dim i As Long
    For i = 1 To NoOfSections
        X = SecName(i)
        If (X = Name And i <> Exception) Then
            IsSecExist = True
            Exit Function
        End If
    Next i
IsSecExist = False
Exit Function
ErrorHandler:

Call subDispErrInfo("IsSectionExist procedure", Err.Number, Err.Description)
End Function


Sub subDispErrInfo(During As String, Number, Description)

On Error GoTo ErrorHandler

    MsgBox "An error occured while " & During & vbCrLf & _
            "Error # " & Number & vbCrLf & _
            "Description: " & Description, , "Error Occured"

Exit Sub
ErrorHandler:
    MsgBox "An error occured while displaying the error information" & vbCrLf & _
            Err.Number & vbCrLf & _
            Err.Description, , "Error Occured"
End Sub



