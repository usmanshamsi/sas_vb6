VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "  STRUCTURAL ANALYSIS SOFTWARE"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10395
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000005&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   60
      ScaleHeight     =   6075
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   60
      Width           =   6855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu menuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu menuFileSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menuAdd 
      Caption         =   "&Add"
      Begin VB.Menu menuAddNode 
         Caption         =   "Add &Node"
      End
      Begin VB.Menu menuAddElement 
         Caption         =   "Add &Element"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menuAddMaterial 
         Caption         =   "Add &Material"
      End
      Begin VB.Menu menuAddSection 
         Caption         =   "Add &Section"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu menuAddRestraint 
         Caption         =   "Apply &Restraint"
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu menuEditDeleteNode 
         Caption         =   "&Node"
         Begin VB.Menu menueditNode 
            Caption         =   "&Edit Node"
         End
         Begin VB.Menu menuDeleteNode 
            Caption         =   "&Delete Node"
         End
      End
      Begin VB.Menu menuEditDeleteElement 
         Caption         =   "&Element"
         Begin VB.Menu menuEditElement 
            Caption         =   "Edit &Element"
         End
         Begin VB.Menu menuDeleteElement 
            Caption         =   "&Delete Element"
         End
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu menuEditMaterial 
         Caption         =   "Edit &Material"
      End
      Begin VB.Menu menuEditSection 
         Caption         =   "Edit &Section"
      End
   End
   Begin VB.Menu menuLoad 
      Caption         =   "&Load"
      Begin VB.Menu menuAsgnNodalLoad 
         Caption         =   "Assign &Nodal Load"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu menuAsgnElemLoad 
         Caption         =   "Assign &Element Load"
      End
   End
   Begin VB.Menu menuAnalyze 
      Caption         =   "A&nalyze"
      Begin VB.Menu menuRunAnalysis 
         Caption         =   "&Run Analysis"
         Shortcut        =   {F5}
      End
      Begin VB.Menu menuDisAnalysis 
         Caption         =   "Discard Analysis"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu menuWriteAnalysisFile 
         Caption         =   "&Write Analysis To File"
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "&View"
      Begin VB.Menu menushowDS 
         Caption         =   "&Deformed Shape"
      End
      Begin VB.Menu menuUndeformedShape 
         Caption         =   "Undeformed Shape"
      End
      Begin VB.Menu menuShowLoads 
         Caption         =   "Show Loads"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep20 
         Caption         =   "-"
      End
      Begin VB.Menu menushowBMD 
         Caption         =   "&Bending Moment Diagram"
      End
      Begin VB.Menu menuShowSF 
         Caption         =   "&Shear Force Diagram"
      End
      Begin VB.Menu menuShowAxial 
         Caption         =   "&Axial Force Diagram"
      End
      Begin VB.Menu SPE 
         Caption         =   "-"
      End
      Begin VB.Menu menuShowR 
         Caption         =   "Display Reactions"
      End
      Begin VB.Menu menuShowDef 
         Caption         =   "Display Deformations"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit

Private Sub Form_Load()
'Me.Picture1.Print "WELCOME TO SAS"
End Sub

Private Sub Form_Resize()
    
    If frmMain.Width > 220 And frmMain.Height > 900 Then
       
        Picture1.Width = Me.Width - 220
        Picture1.Height = Me.Height - 900
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim resp
    If ChangesSaved = False Then
        resp = MsgBox("Do you want to save changes you have made to " & FileTitle, vbQuestion Or vbYesNoCancel, "")
        
        If resp = vbCancel Then
            Cancel = 1
            Exit Sub
        ElseIf resp = vbYes Then
            Call menuFileSave_Click
        End If
    End If
        
End Sub

Private Sub menuAbout_Click()
    Load frmAbout
    frmAbout.Show (1)
End Sub

Private Sub menuAddNode_Click()
    Load frmAddNode
    With frmAddNode
        .Caption = "ADD NODE"
        .Label5.Visible = False
        .txtXCoOr.Text = 0
        .txtYCoOr.Text = 0
        .txtNodeNo.Enabled = False
        .txtNodeNo.Text = NoofNodes + 1
        .Show (1)
    End With
End Sub


Private Sub menuAddMaterial_Click()
Load frmAddMaterial
With frmAddMaterial
    .Caption = "ADD MATERIAL"
    .txtMatCoeffTher.Text = 0
    .txtMatID.Enabled = False
    .txtMatID.Text = NoOfMaterials + 1
    .Show (1)
End With
End Sub

Private Sub menuAddElement_Click()
Load frmAddElement
With frmAddElement
    .Caption = "ADD ELEMENT"
    .txtElementNo.Enabled = False
    .Label5.Visible = False
    .txtElementNo.Text = NoOfElements + 1
Dim i As Long
    For i = 1 To NoOfSections
        .cmbSelSec.AddItem SecName(i)
    Next i
    .cmbSelSec.ListIndex = 1 - 1
    .Show (1)
End With

End Sub


Private Sub menuAddRestraint_Click()
    Load frmAddRestraint
    frmAddRestraint.Show (1)
End Sub

Private Sub menuAddSection_Click()
Load frmAddSection
With frmAddSection
    .Caption = "ADD SECTION"
    .txtSecID.Enabled = False
    .txtSecID.Text = NoOfSections + 1
Dim i As Long
    For i = 1 To NoOfMaterials
        .cmbSelMat.AddItem MaterialName(i)
    Next i
    .cmbSelMat.ListIndex = 1 - 1
    .Show (1)
End With
End Sub

Private Sub menuAsgnElemLoad_Click()
    Load frmAsgnElemLoad
    With frmAsgnElemLoad
        .txtElemNo.Text = NoOfElements
        .txtUDL1.Text = 0
        .txtUDL2.Text = 0
        .Show (1)
    End With
End Sub

Private Sub menuAsgnNodalLoad_Click()
    Load frmAsgnNodalLoad
    With frmAsgnNodalLoad
        .txtNodeNo.Text = NoofNodes
        .txtXForce.Text = 0
        .txtYForce.Text = 0
        .txtZForce.Text = 0
        .txtXMom.Text = 0
        .txtYMom.Text = 0
        .txtZMom.Text = 0
        .Show (1)
    End With
End Sub

Private Sub menuDeleteElement_Click()
Load frmAddElement
With frmAddElement
    .txtElementNo.Enabled = True
    .txtElementNo.Text = ""
Dim i As Long
    For i = 1 To NoOfSections
        .cmbSelSec.AddItem SecName(i)
    Next i
    .Caption = "DELETE ELEMENT"
    .btnAddElement.Caption = "&Delete"
    .cmbSelSec.ListIndex = 1 - 1
    .Show (1)
End With
End Sub

Private Sub menuDeleteNode_Click()
    Load frmAddNode
    With frmAddNode
        .Caption = "DELETE NODE"
        .btnAdd.Caption = "&Delete"
        .txtNodeNo.Enabled = True
        .txtNodeNo.Text = NoofNodes
        .Show (1)
    End With
End Sub

Private Sub menuDisAnalysis_Click()
Analyzed = False
PlotBMD = False
PlotSF = False
menushowBMD.Checked = False
menuShowSF.Checked = False
    
subPlotStr
End Sub

Private Sub menuEditElement_Click()
Load frmAddElement
With frmAddElement
    .txtElementNo.Enabled = True
    .txtElementNo.Text = ""
Dim i As Long
    For i = 1 To NoOfSections
        .cmbSelSec.AddItem SecName(i)
    Next i
    .Caption = "EDIT ELEMENT"
    .btnAddElement.Caption = "&Modify"
    .cmbSelSec.ListIndex = 1 - 1
    .Show (1)
End With
End Sub

Private Sub menuEditMaterial_Click()
Load frmAddMaterial
With frmAddMaterial
    .Caption = "EDIT MATERIAL"
    .btnAddMat.Caption = "&Modify"
    .txtMatID.Enabled = True
    .Show (1)
End With
End Sub

Private Sub menuEditNode_Click()
    Load frmAddNode
    With frmAddNode
        .Caption = "EDIT NODE"
        .btnAdd.Caption = "&Modify"
        .txtNodeNo.Enabled = True
        .txtNodeNo.Text = NoofNodes
        .Show (1)
    End With
End Sub

Private Sub menuEditSection_Click()
Load frmAddSection
With frmAddSection
    .Caption = "EDIT SECTION"
    .btnAddSec.Caption = "&Modify"
    .txtSecID.Enabled = True
Dim i As Long
    For i = 1 To NoOfMaterials
        .cmbSelMat.AddItem MaterialName(i)
    Next i
    .cmbSelMat.ListIndex = 1 - 1
    .Show (1)
End With
End Sub



Private Sub menuFileExit_Click()
Unload Me
End Sub

Private Sub menuFileNew_Click()
    Dim resp
    If ChangesSaved = False Then
        resp = MsgBox("Do you want to save changes you have made to " & FileTitle, vbQuestion Or vbYesNoCancel, "")
        
        If resp = vbCancel Then
            Exit Sub
        ElseIf resp = vbYes Then
            Call menuFileSave_Click
        End If
    End If
    LoadDefVal
    
    subPlotStr
    
End Sub

Private Sub menuFileOpen_Click()
On Error GoTo NoFile
    CommonDialog1.Filter = "SAS Files|*.sas|Text File|*.txt|All Files|*.*"
    CommonDialog1.DefaultExt = "*.sas"
    CommonDialog1.ShowOpen
    filename = CommonDialog1.filename
    FileTitle = CommonDialog1.FileTitle

    Call subOpenFile(filename)
      
    
    'MsgBox FileTitle & " is successfully opened", , "File Opened"
    
    Call subPlotStr
Exit Sub

NoFile:
If Err.Number = 32755 Then
    Exit Sub    'operation canceled
Else
    MsgBox Err.Number & vbCrLf & Err.Description
    
End If

End Sub

Private Sub menuFileSave_Click()
On Error GoTo NoFile

    If FileSaved = False Then
        Call menuFileSaveAs_Click
    Else
        Call subSaveFile(filename)
    End If

Exit Sub

NoFile:
If Err.Number = 32755 Then
    Exit Sub    'operation canceled
Else
    Call subDispErrInfo("saving file", Err.Number, Err.Description)
End If
End Sub

Private Sub menuFileSaveAs_Click()
On Error GoTo NoFile
    CommonDialog1.Filter = "SAS Files|*.sas|Text File|*.txt|All Files|*.*"
    CommonDialog1.DefaultExt = "*.sas"
    
    CommonDialog1.ShowSave
    filename = CommonDialog1.filename
    

    
    FileTitle = CommonDialog1.FileTitle

    Call subSaveFile(filename)
    
    MsgBox "The Project is successfully save to " & vbCrLf & _
            filename, , "File Saved"

Exit Sub
NoFile:
If Err.Number = 32755 Then
    Exit Sub    'operation canceled
Else
    Call subDispErrInfo("saving file", Err.Number, Err.Description)
End If
End Sub

Private Sub menuRunAnalysis_Click()
    Call subSolve
End Sub


Private Sub menuShowAxial_Click()
On Error GoTo ERRORHANDLE
'MsgBox "This feature is currently unavailable"
'Exit Sub
If Analyzed = False Then
    MsgBox "PLEASE RUN ANALYSIS..."
    Exit Sub
End If


    PlotAxial = True
    PlotBMD = False
    PlotSF = False
    PlotDeflectedShape = False


AxialScale = InputBox("Current Axial Force Scale = " & AxialScale _
                            & vbCrLf & "Enter New Axial Force Scale", , AxialScale)

Call subCalcAxialShapes

Call subPlotStr

ERRORHANDLE:
End Sub

Private Sub menushowBMD_Click()
On Error GoTo ERRORHANDLE
'MsgBox "This feature is currently unavailable"
'Exit Sub
If Analyzed = False Then
    MsgBox "PLEASE RUN ANALYSIS..."
    Exit Sub
End If


    PlotBMD = True
    PlotAxial = False
    PlotSF = False
    PlotDeflectedShape = False


BMDScale = InputBox("Current Bending Moment Scale = " & BMDScale _
                            & vbCrLf & "Enter New Bending Moment Scale", , BMDScale)

Call subCalcBMDShapes

Call subPlotStr

ERRORHANDLE:
End Sub

Private Sub menuShowDef_Click()
If Analyzed = False Then
    MsgBox "PLEASE RUN ANALYSIS..."
    Exit Sub
End If
PlotBMD = False
PlotDeflectedShape = True
PlotSF = False
PlotAxial = False
subPlotStr

Dim txt As String, i As Long
    txt = "Node No.    Horizontal           Vertical                Rotation"
    txt = txt & vbCrLf & _
          "-----------------------------------------------------------------------------"
    For i = 1 To NoofNodes
            txt = txt & vbCrLf & _
          Format(i, "000") & "             " & _
          Format(DispVector(3 * (i - 1) + 1, 1), "000.00000000") & "    " & _
          Format(DispVector(3 * (i - 1) + 2, 1), "000.00000000") & "    " & _
          Format(DispVector(3 * (i - 1) + 3, 1), "000.00000000")
    Next i
    MsgBox txt, , "Nodal Deformations"
End Sub

Private Sub menushowDS_Click()
On Error GoTo ERRORHANDLE
If Analyzed = False Then
    MsgBox "PLEASE RUN ANALYSIS..."
    Exit Sub
End If

'MsgBox "This feature is currently unavailable"

PlotAxial = False
PlotBMD = False
PlotSF = False
PlotDeflectedShape = True
 DeflectionScale = InputBox("Current Deflection Scale = " & DeflectionScale _
                            & vbCrLf & "Enter New Deflection Scale", , DeflectionScale)
Call subCalcDefCoor
Call subCalcDefShapes

Call subPlotStr

ERRORHANDLE:
End Sub

Private Sub menuShowLoads_Click()
menuShowLoads.Checked = Not (menuShowLoads.Checked)
subPlotStr
End Sub

Private Sub menuShowR_Click()
If Analyzed = False Then
    MsgBox "PLEASE RUN ANALYSIS..."
    Exit Sub
End If
PlotBMD = False
PlotDeflectedShape = False
PlotSF = False
menushowBMD.Checked = False
menuShowSF.Checked = False
PlotAxial = False
subPlotStr

Dim txt As String, i As Long, Line As Boolean
Line = False
    For i = 1 To NoofNodes
            If Line = True Then
                txt = txt & vbCrLf & vbCrLf
                Line = False
            End If
            If RestVector(3 * (i - 1) + 1, 1) = 1 Then
                txt = txt & "H" & i & " = " & Format(ForceVector(3 * (i - 1) + 1, 1), "0.0000") & "    "
                Line = True
            End If
            If RestVector(3 * (i - 1) + 2, 1) = 1 Then
                txt = txt & "V" & i & " = " & Format(ForceVector(3 * (i - 1) + 2, 1), "0.0000") & "    "
                Line = True
            End If
            If RestVector(3 * (i - 1) + 3, 1) = 1 Then
                txt = txt & "M" & i & " = " & Format(ForceVector(3 * (i - 1) + 3, 1), "0.0000") & "    "
                Line = True
            End If

            
    Next i
    MsgBox txt, , "REACTIONS"
End Sub

Private Sub menuShowSF_Click()
On Error GoTo ERRORHANDLE
'MsgBox "This feature is currently unavailable"

'Exit Sub
If Analyzed = False Then
    MsgBox "PLEASE RUN ANALYSIS..."
    Exit Sub
End If


    PlotAxial = False
    PlotBMD = False
    PlotSF = True
    PlotDeflectedShape = False



SFScale = InputBox("Current Shear Force Scale = " & SFScale _
                            & vbCrLf & "Enter New Shear Force Scale", , SFScale)

Call subCalcSFShapes

Call subPlotStr

ERRORHANDLE:
End Sub

Private Sub menuUndeformedShape_Click()

PlotBMD = False
PlotDeflectedShape = False
PlotSF = False
PlotAxial = False

subPlotStr
End Sub

Private Sub menuWriteAnalysisFile_Click()
On Error GoTo NoFile
If Analyzed = False Then
    MsgBox "PLEASE RUN ANALYSIS..."
    Exit Sub
End If
CommonDialog1.Filter = "Text File|*.txt|All Files|*.*"
CommonDialog1.DefaultExt = "*.txt"

Dim AnalysisFileName As String
Dim temp() As Double, i As Long, j As Long

CommonDialog1.ShowSave
AnalysisFileName = CommonDialog1.filename
    
Open AnalysisFileName For Output Access Write As #1

'START WRITING FILE, WRITE FILENAME
'----------------------------------
    Print #1, filename
    Print #1,
'----------------------------------

'MATERIAL DATA
'----------------------------------
    Print #1, "Material Data"
    Print #1, "============="
    Print #1,
    Print #1, "Number of materials defined = ", NoOfMaterials
    Print #1,
    Print #1, "Material ID", "Name", " E ", " G ", "Alpha"
    Print #1, "-----------", "----", "---", "---", "-----"
    For i = 1 To NoOfMaterials
        Print #1, i, MaterialName(i), MatElasMod(i), MatShearMod(i), MatCoeffTher(i)
    Next i
    Print #1,
'----------------------------------

'SECTION DATA
'----------------------------------
    Print #1,
    Print #1, "Section's Data"
    Print #1, "=============="
    Print #1,
    Print #1, "Number of Sections defined = ", NoOfSections
    Print #1,
    Print #1, "Section ID", "Name", "Material", "Area", "I-x", "I-y", " J ", "Color"
    Print #1, "----------", "----", "--------", "----", "---", "---", "---", "-----"
    For i = 1 To NoOfSections
        Print #1, i, SecName(i), SecMat(i), SecArea(i), SecIx(i), SecIy(i), SecJ(i), SecColor(i)
    Next i
    Print #1,
'----------------------------------

'NODAL DATA
'----------------------------------
    Print #1,
    Print #1, "Nodal Data"
    Print #1, "=========="
    Print #1,
    Print #1, "Number of Nodes =", NoofNodes
    Print #1,
    Print #1, "Node #", "X-Coordinate", "Y-Coordinate"
    Print #1, "------", "------------", "------------"
    For i = 1 To NoofNodes
        Print #1, i, Xcoor(i), Ycoor(i)
    Next i
    Print #1,
'----------------------------------


'ELEMENT DATA
'----------------------------------
    Print #1,
    Print #1, "Element Data"
    Print #1, "============"
    Print #1,
    Print #1, "Number of Elements =", NoOfElements
    Print #1,
    Print #1, "Element#", "First Node", "Second Node", "Section"
    Print #1, "--------", "----------", "-----------", "-------"
    For i = 1 To NoOfElements
        Print #1, i, Endi(i), Endj(i), AsgnSec(i)
    Next i
    Print #1,
'----------------------------------

'NODAL FORCES
'----------------------------------
    Print #1,
    Print #1, "Nodal Forces"
    Print #1, "============"
    Print #1,
    Print #1, "Node #", "Force Along X", "Force Along Y", "Moment About Z"
    Print #1, "------", "-------------", "-------------", "--------------"
    For i = 1 To NoofNodes
        Print #1, i, XForce(i), YForce(i), ZMom(i)
    Next i
    Print #1,
'----------------------------------

'ELEMENT FORCES
'----------------------------------
    Print #1,
    Print #1, "Element Forces"
    Print #1, "=============="
    Print #1,
    Print #1, "Element#", "UDL at Node 1", "UDL at Node 2", "UDAL at Node 1", "UDAL at Node 2"
    Print #1, "--------", "-------------", "-------------"
    For i = 1 To NoOfElements
        Print #1, i, ElemLoadi(i), ElemLoadj(i), ElemALoadi(i), ElemALoadj(i)
    Next i
    Print #1,
'----------------------------------

'RESTRAINTS
'----------------------------------
    Print #1,
    Print #1, "Restraints"
    Print #1, "=========="
    Print #1,
    Print #1, "Node #", "Trans-X", "Trans-Y", "Rotation-Z"
    Print #1, "------", "-------", "-------", "----------"
    For i = 1 To NoofNodes
        Print #1, i, TxRest(i), TyRest(i), RzRest(i)
    Next i
    Print #1,
'----------------------------------

If Analyzed = False Then
    Close #1
    Exit Sub
End If

'ELELMENTS' LOCAL, GLOBAL AND TRANSFORMATION MATRICES
'-----------------------------------------------------
    Print #1,
    Print #1, "ELEMENTS' MATRICES"
    Print #1, "=================="
    Print #1,
    For i = 1 To NoOfElements
        Call GetElemLocalStiff(i, temp())
        Print #1, "Local Stiffness Matrix for element # " & Str(i)
        Print #1, "------------------------------------------"
        Call subWriteMatrix(temp(), "0.0000", 1)
        
        Call GetTrans(i, temp())
        Print #1, "Transformation Matrix for element # " & Str(i)
        Print #1, "------------------------------------------"
        Call subWriteMatrix(temp(), "0.0000", 1)
        
        Call GetElemGlobalStiff(i, temp())
        Print #1, "Global Stiffness Matrix for element # " & Str(i)
        Print #1, "------------------------------------------"
        Call subWriteMatrix(temp(), "0.0000", 1)
        Print #1,
    Next i
    
'-----------------------------------------------------

'SYSTEM STIFFNESS MATRICES
'-----------------------------------------------------
    Print #1, "Global Stiffness Matrix for the whole Structure"
    Print #1, "==============================================="
    Call subWriteMatrix(SysStiff(), "0.0000", 1)
    Print #1,
    
    Print #1, "Reduced Global Stiffness Matrix for the whole Structure"
    Print #1, "======================================================="
    Call subWriteMatrix(RedSysStiff(), "0.0000", 1)
    Print #1,
    
'-----------------------------------------------------

'DISPLACEMENTS
'-----------------------------------------------------
    Print #1, "Displacements"
    Print #1, "============="
    Print #1,
    Print #1, "Node #", "Along X", "Along Y", "About Z"
    Print #1, "------", "-------", "-------", "-------"
    For i = 1 To NoofNodes
        Print #1, i, Format(DispVector(3 * (i - 1) + 1, 1), "0.00000000"), _
            Format(DispVector(3 * (i - 1) + 2, 1), "0.00000000"), _
            Format(DispVector(3 * (i - 1) + 3, 1), "0.00000000")
    Next i
    Print #1,

'-----------------------------------------------------

'FORCES
'-----------------------------------------------------
    Print #1, "Nodal Forces"
    Print #1, "============="
    Print #1,
    Print #1, "Node #", "Along X", "Along Y", "About Z"
    Print #1, "------", "-------", "-------", "-------"
    For i = 1 To NoofNodes
        Print #1, i, Format(ForceVector(3 * (i - 1) + 1, 1), "0.0000"), _
            Format(ForceVector(3 * (i - 1) + 2, 1), "0.0000"), _
            Format(ForceVector(3 * (i - 1) + 3, 1), "0.0000")
    Next i
    Print #1,

'-----------------------------------------------------

'INTERNAL FORCES
'----------------
    Dim space
    space = 26
    Print #1, "INTERNAL FORCES"
    Print #1, "==============="
    Print #1,
    
    For i = 1 To NoOfElements
        Print #1, "ELEMENT NO. " & i
        Print #1, "---------------"
        Print #1,
        Print #1, "DISTANCE", _
                Tab(space); "SHEAR_FORCE", _
                Tab(2 * space); "BENDING_MOMENT", _
                Tab(3 * space); "AXIAL-FORCE", _
                Tab(4 * space); "DEFLECTION", _
                Tab(5 * space); "SLOPE"
                
        Print #1, "--------", _
                Tab(space); "-----------", _
                Tab(2 * space); "--------------", _
                Tab(3 * space); "-----------", _
                Tab(4 * space); "----------", _
                Tab(5 * space); "-----"
        For j = 0 To Mesh
            Print #1, Format(GetElemLength(i) / Mesh * j, "0.000"), _
                Tab(space); Format(ShearForce(i, j), "0.0000"), _
                Tab(2 * space); Format(BendingMoment(i, j), "0.0000"), _
                Tab(3 * space); Format(AxialForce(i, j), "0.0000"), _
                Tab(4 * space); Format(Deflection(i, j), "0.00000000"), _
                Tab(5 * space); Format(Slope(i, j), "0.00000000")
        Next j
        Print #1,
    Next i
    Print #1,
    


Close #1

Exit Sub

NoFile:
If Err.Number = 32755 Then
    Exit Sub    'operation canceled
Else
    Call subDispErrInfo("saving file", Err.Number, Err.Description)
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.Caption = FileTitle & " (" & Format(X, "#0.0000") & " , " & _
                  Format(YMirror(Y), "#0.0000") & " )"
                  
If PlotBMD = True Then frmMain.Caption = frmMain.Caption + " Bending Moment"

If PlotSF = True Then frmMain.Caption = frmMain.Caption + " Shear Force"

If PlotDeflectedShape = True Then frmMain.Caption = frmMain.Caption + " Deflected Shape"

If PlotAxial = True Then frmMain.Caption = frmMain.Caption + " Axial Force"
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MsgBox X & vbCrLf & Y & vbCrLf & X * PlotScale & vbCrLf & Y * PlotScale


End Sub

Private Sub Picture1_Paint()
subPlotStr
DoEvents
End Sub

Private Sub Picture1_Resize()

Dim T1 As Boolean, T2 As Boolean, T3 As Boolean, T4 As Boolean
T1 = PlotDeflectedShape
T2 = PlotBMD
T3 = PlotSF
T4 = PlotAxial

PlotDeflectedShape = False
PlotBMD = False
PlotSF = False
PlotAxial = False

subPlotStr
DoEvents

PlotDeflectedShape = T1
PlotBMD = T2
PlotSF = T3
PlotAxial = T4

subPlotStr
DoEvents

End Sub
