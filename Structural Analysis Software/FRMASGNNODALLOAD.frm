VERSION 5.00
Begin VB.Form frmAsgnNodalLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Nodal Loads"
   ClientHeight    =   3345
   ClientLeft      =   3945
   ClientTop       =   4365
   ClientWidth     =   9060
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
   ScaleHeight     =   3345
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnShowLoad 
      Caption         =   "&Show Existing Loads"
      Height          =   495
      Left            =   6600
      TabIndex        =   20
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton btnDeleteLoad 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton btnReplaceLoad 
      Caption         =   "&Replace"
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtNodeNo 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton btnAddLoad 
      Caption         =   "&Add"
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   6600
      TabIndex        =   18
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Moment"
      Height          =   2415
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   3015
      Begin VB.TextBox txtYMom 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtZMom 
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtXMom 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "About Z:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "About Y:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label55 
         Caption         =   "About X:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Force"
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3015
      Begin VB.TextBox txtZForce 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtYForce 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtXForce 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Along Z:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Along Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Along X:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   300
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "&Node No:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   300
      Width           =   1095
   End
End
Attribute VB_Name = "frmAsgnNodalLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit

Private Sub btnAddLoad_Click()
            
    'CHECK IF THE BOXES ARE FILLED
        If (txtNodeNo.Text) = "" Then
                MsgBox "Enter Node Number"
                Exit Sub
        End If
    'ADD THE NODAL FORCES
        Call subAddNodalLoad(Val(txtNodeNo.Text), _
                                Val(txtXForce.Text), _
                                Val(txtYForce.Text), _
                                Val(txtZMom.Text))

End Sub


Private Sub btnClose_Click()
    frmAsgnNodalLoad.Hide
    Unload frmAsgnNodalLoad
End Sub

Private Sub btnDeleteLoad_Click()
            
    'CHECK IF THE BOXES ARE FILLED
        If (txtNodeNo.Text) = "" Then
                MsgBox "Enter Node Number"
                Exit Sub
        End If
    'ADD THE NODAL FORCES
        Call subReplaceNodalLoad(Val(txtNodeNo.Text), 0, 0, 0)

End Sub

Private Sub btnReplaceLoad_Click()
            
    'CHECK IF THE BOXES ARE FILLED
        If (txtNodeNo.Text) = "" Then
                MsgBox "Enter Node Number"
                Exit Sub
        End If
    'ADD THE NODAL FORCES
        Call subReplaceNodalLoad(Val(txtNodeNo.Text), _
                                Val(txtXForce.Text), _
                                Val(txtYForce.Text), _
                                Val(txtZMom.Text))


End Sub

Private Sub btnShowLoad_Click()
Dim Value As Long
    Value = Val(txtNodeNo.Text)
    If Value <= NoofNodes And Value > 0 Then
        txtXForce.Text = XForce(Value)
        txtYForce.Text = YForce(Value)
        txtZMom.Text = ZMom(Value)
    End If
End Sub

Private Sub txtNodeNo_Change()
Dim Value As Long
    Value = Val(txtNodeNo.Text)
    
    If Value <= NoofNodes And Value > 0 Then
        btnAddLoad.Enabled = True
        btnShowLoad.Enabled = True
        btnReplaceLoad.Enabled = True
        btnDeleteLoad.Enabled = True
        Label5.Caption = ""
    Else
        Label5.Caption = "(Invalid Node Number...)"
        btnAddLoad.Enabled = False
        btnShowLoad.Enabled = False
        btnReplaceLoad.Enabled = False
        btnDeleteLoad.Enabled = False
    End If


End Sub

Private Sub txtNodeNo_GotFocus()
txtNodeNo.SelStart = 0
txtNodeNo.SelLength = Len(txtNodeNo.Text)
End Sub


Private Sub txtXForce_GotFocus()
txtXForce.SelStart = 0
txtXForce.SelLength = Len(txtXForce.Text)
End Sub


Private Sub txtXMom_GotFocus()
txtXMom.SelStart = 0
txtXMom.SelLength = Len(txtXMom.Text)
End Sub


Private Sub txtYForce_GotFocus()
txtYForce.SelStart = 0
txtYForce.SelLength = Len(txtYForce.Text)
End Sub


Private Sub txtYMom_GotFocus()
txtYMom.SelStart = 0
txtYMom.SelLength = Len(txtYMom.Text)
End Sub


Private Sub txtZForce_GotFocus()
txtZForce.SelStart = 0
txtZForce.SelLength = Len(txtZForce.Text)
End Sub



Private Sub txtZMom_GotFocus()
txtZMom.SelStart = 0
txtZMom.SelLength = Len(txtZMom.Text)
End Sub
