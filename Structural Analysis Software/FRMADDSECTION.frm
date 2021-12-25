VERSION 5.00
Begin VB.Form frmAddSection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Section"
   ClientHeight    =   4590
   ClientLeft      =   6450
   ClientTop       =   5985
   ClientWidth     =   5700
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSelectColor 
      Caption         =   "..."
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox picSecColor 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4440
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox cmbSelMat 
      Height          =   330
      ItemData        =   "frmAddSection.frx":0000
      Left            =   3000
      List            =   "frmAddSection.frx":0002
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   500
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
      Width           =   1700
   End
   Begin VB.CommandButton btnAddSec 
      Caption         =   "&Add"
      Height          =   500
      Left            =   3360
      TabIndex        =   9
      Top             =   3120
      Width           =   1700
   End
   Begin VB.Frame SelMat 
      Caption         =   "Material Property"
      Height          =   2535
      Left            =   7800
      TabIndex        =   18
      Top             =   2040
      Width           =   2895
      Begin VB.TextBox txtSecElasMod 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtSecShearMod 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtSecCoeffTher 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Alpha:"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "G:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "E:"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame framesection 
      Caption         =   "Geometric Properties"
      Height          =   2775
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
      Begin VB.TextBox txtSecJ 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtSecIy 
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtSecIx 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtSecArea 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Ix:"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "J:"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Iy:"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Area:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.TextBox txtSecName 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtSecID 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Display Color:"
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label lblMaterialInfo 
      Height          =   735
      Left            =   3000
      TabIndex        =   25
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Material:"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Section ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit

Private Sub btnAddSec_Click()
    'CHECK IF THE FORM IS FILLED OR ANY VALUE IS MISSING
        If txtSecName.Text = "" Then
            MsgBox "Enter Section Name"
            Exit Sub
        End If
        If cmbSelMat.Text = "" Then
            MsgBox "Please Assing Material"
            Exit Sub
        End If
        
        If txtSecArea.Text = "" Then
            MsgBox "Section Property Missing"
            Exit Sub
        End If
        
        If txtSecIx.Text = "" Then
            MsgBox "Section Property Missing"
            Exit Sub
        End If
        If txtSecIy.Text = "" Then
            MsgBox "Section Property Missing"
            Exit Sub
        End If
        If txtSecJ.Text = "" Then
            MsgBox "Section Property Missing"
            Exit Sub
        End If
    
    'ADD OR EDIT ACCORDINGLY
    
    Select Case btnAddSec.Caption
    
        Case "&Add"
            Call subAddSection(txtSecName.Text, _
                                picSecColor.BackColor, _
                                cmbSelMat.ListIndex + 1, _
                                Val(txtSecArea.Text), _
                                Val(txtSecIx.Text), _
                                Val(txtSecIy.Text), _
                                Val(txtSecJ.Text))
                                
            txtSecID.Text = NoOfSections + 1
        
        Case "&Modify"
            Call subEditSection(Val(txtSecID.Text), _
                                txtSecName.Text, _
                                picSecColor.BackColor, _
                                cmbSelMat.ListIndex + 1, _
                                Val(txtSecArea.Text), _
                                Val(txtSecIx.Text), _
                                Val(txtSecIy.Text), _
                                Val(txtSecJ.Text))
    End Select


End Sub

Private Sub btnClose_Click()
frmAddSection.Hide
Unload frmAddSection
End Sub

Private Sub btnSelectColor_Click()
frmMain.CommonDialog1.CancelError = True
On Error GoTo ErrorHandler
    
    frmMain.CommonDialog1.ShowColor
    picSecColor.BackColor = frmMain.CommonDialog1.Color
    
Exit Sub

ErrorHandler:
If Err.Number = 32755 Then Exit Sub
Call subDispErrInfo("selecting color", Err.Number, Err.Description)
End Sub

Private Sub cmbSelMat_Click()

    'MUQEET'S OPTION
        txtSecElasMod.Text = MatElasMod(cmbSelMat.ListIndex + 1)
        txtSecShearMod.Text = MatShearMod(cmbSelMat.ListIndex + 1)
        txtSecCoeffTher.Text = MatCoeffTher(cmbSelMat.ListIndex + 1)
    
    'USMAN'S OPTION
    Dim txt As String
        txt = "E = " & MatElasMod(cmbSelMat.ListIndex + 1) & vbCrLf
        txt = txt & "G = " & MatShearMod(cmbSelMat.ListIndex + 1) & vbCrLf
        txt = txt & "Alpha = " & MatCoeffTher(cmbSelMat.ListIndex + 1)
        lblMaterialInfo.Caption = txt
        
End Sub



Private Sub txtSecArea_GotFocus()
txtSecArea.SelStart = 0
txtSecArea.SelLength = Len(txtSecArea.Text)
End Sub




Private Sub txtSecCoeffTher_GotFocus()
txtSecCoeffTher.SelStart = 0
txtSecCoeffTher.SelLength = Len(txtSecCoeffTher.Text)
End Sub



Private Sub txtSecElasMod_GotFocus()
txtSecElasMod.SelStart = 0
txtSecElasMod.SelLength = Len(txtSecElasMod.Text)
End Sub

Private Sub txtSecID_Change()
Dim Value As Long
    If btnAddSec.Caption = "&Add" Then Exit Sub
    Value = Val(txtSecID.Text)
    
    txtSecArea.Text = ""
    txtSecIx.Text = ""
    txtSecIy.Text = ""
    txtSecJ.Text = ""
    txtSecName.Text = ""
    cmbSelMat.ListIndex = 0
    picSecColor.BackColor = vbWhite
    btnAddSec.Enabled = False
    
    If Value <= NoOfSections And Value > 0 Then
        txtSecArea.Text = SecArea(Value)
        txtSecIx.Text = SecIx(Value)
        txtSecIy.Text = SecIy(Value)
        txtSecJ.Text = SecJ(Value)
        txtSecName.Text = SecName(Value)
        cmbSelMat.ListIndex = SecMat(Value) - 1
        picSecColor.BackColor = SecColor(Value)
        btnAddSec.Enabled = True
        
    End If


End Sub

Private Sub txtSecID_GotFocus()
txtSecID.SelStart = 0
txtSecID.SelLength = Len(txtSecID.Text)
End Sub

Private Sub txtSecIx_GotFocus()
txtSecIx.SelStart = 0
txtSecIx.SelLength = Len(txtSecIx.Text)
End Sub
Private Sub txtSecIy_GotFocus()
txtSecIy.SelStart = 0
txtSecIy.SelLength = Len(txtSecIy.Text)
End Sub


Private Sub txtSecJ_GotFocus()
txtSecJ.SelStart = 0
txtSecJ.SelLength = Len(txtSecJ.Text)
End Sub

Private Sub txtSecName_GotFocus()
txtSecName.SelStart = 0
txtSecName.SelLength = Len(txtSecName.Text)
End Sub


Private Sub txtSecShearMod_GotFocus()
txtSecShearMod.SelStart = 0
txtSecShearMod.SelLength = Len(txtSecShearMod.Text)
End Sub
