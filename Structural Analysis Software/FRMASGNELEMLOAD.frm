VERSION 5.00
Begin VB.Form frmAsgnElemLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Element Load"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
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
   ScaleHeight     =   3360
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Load per unit length:"
      Height          =   2175
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   6375
      Begin VB.TextBox txtUDAL2 
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtUDAL1 
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtUDL2 
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtUDL1 
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Horizontal"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblNode2 
         Caption         =   "At Node # 22222222:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1500
         Width           =   3495
      End
      Begin VB.Label lblNode1 
         Caption         =   "At Node #   :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   900
         Width           =   3495
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton btnAddLoad 
      Caption         =   "&Add"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtElemNo 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnReplaceLoad 
      Caption         =   "&Replace"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton btnDeleteLoad 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton btnShowLoad 
      Caption         =   "&Show Existing Loads"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "&Element No:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "frmAsgnElemLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit
Private Sub btnAddLoad_Click()
            
    'CHECK IF THE BOXES ARE FILLED
        If (txtElemNo.Text) = "" Then
                MsgBox "Enter Element Number"
                Exit Sub
        End If
    'ADD THE element FORCES
        Call subAddElemLoad(Val(txtElemNo.Text), Val(txtUDL1), Val(txtUDL2), Val(txtUDAL1), Val(txtUDAL2))

End Sub


Private Sub btnClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub btnDeleteLoad_Click()
            
    'CHECK IF THE BOXES ARE FILLED
        If (txtElemNo.Text) = "" Then
                MsgBox "Enter Element Number"
                Exit Sub
        End If
    'ADD THE NODAL FORCES
        Call subReplaceElemLoad(Val(txtElemNo.Text), 0, 0, 0, 0)

End Sub

Private Sub btnReplaceLoad_Click()
            
    'CHECK IF THE BOXES ARE FILLED
        If (txtElemNo.Text) = "" Then
                MsgBox "Enter Element Number"
                Exit Sub
        End If
    'ADD THE element FORCES
        Call subReplaceElemLoad(Val(txtElemNo.Text), Val(txtUDL1), Val(txtUDL2), Val(txtUDAL1), Val(txtUDAL2))

End Sub

Private Sub btnShowLoad_Click()
Dim Value
    Value = Val(txtElemNo.Text)
    If Value <= NoOfElements And Value > 0 Then
        txtUDL1.Text = ElemLoadi(Value)
        txtUDL2.Text = ElemLoadj(Value)
        txtUDAL1.Text = ElemALoadi(Value)
        txtUDAL2.Text = ElemALoadj(Value)
    End If
End Sub

Private Sub txtElemNo_Change()
Dim Value
    Value = Val(txtElemNo.Text)
    
    If Value <= NoOfElements And Value > 0 Then
        btnAddLoad.Enabled = True
        btnShowLoad.Enabled = True
        btnReplaceLoad.Enabled = True
        btnDeleteLoad.Enabled = True
        Label5.Caption = ""
        lblNode1.Caption = "At Node# " & Endi(Value)
        lblNode2.Caption = "At Node# " & Endj(Value)
    Else
        Label5.Caption = "(Invalid Element Number...)"
        btnAddLoad.Enabled = False
        btnShowLoad.Enabled = False
        btnReplaceLoad.Enabled = False
        btnDeleteLoad.Enabled = False
        lblNode1.Caption = "At Node# ..."
        lblNode2.Caption = "At Node# ..."
    End If


End Sub

Private Sub txtElemNo_GotFocus()
txtElemNo.SelStart = 0
txtElemNo.SelLength = Len(txtElemNo.Text)
End Sub



Private Sub txtUDL1_GotFocus()
txtUDL1.SelStart = 0
txtUDL1.SelLength = Len(txtUDL1.Text)
End Sub

Private Sub txtUDL2_GotFocus()
txtUDL2.SelStart = 0
txtUDL2.SelLength = Len(txtUDL2.Text)
End Sub

