VERSION 5.00
Begin VB.Form frmAddElement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Element"
   ClientHeight    =   2385
   ClientLeft      =   6570
   ClientTop       =   4710
   ClientWidth     =   7860
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
   ScaleHeight     =   2385
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSelSec 
      Height          =   330
      Left            =   4320
      TabIndex        =   2
      Top             =   322
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "End Nodes"
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   5535
      Begin VB.TextBox txtNodeI 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   442
         Width           =   1215
      End
      Begin VB.TextBox txtNodeJ 
         Height          =   330
         Left            =   4080
         TabIndex        =   4
         Top             =   442
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Initial Node:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Final Node:"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton btnAddElement 
      Caption         =   "&Add"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtElementNo 
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Top             =   322
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "(Element do not exist)"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Section:"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Element No:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit
Private Sub btnAddElement_Click()

'CHECK IF THE TEXTBOXES ARE FILLED
'---------------------------------
    If txtNodeI.Text = "" Then
        MsgBox "Enter Initial Node"
        Exit Sub
    End If
    
    If txtNodeJ.Text = "" Then
        MsgBox "Enter Final Node"
        Exit Sub
    End If
    
    
'ADD OR EDIT ELEMENT ACCORDINGLY
'----------------------------------

Select Case btnAddElement.Caption

    Case "&Add"
        Call subAddElement(Val(txtNodeI.Text), Val(txtNodeJ.Text), cmbSelSec.ListIndex + 1)
        txtElementNo.Text = NoOfElements + 1
        Call subPlotStr
    
    Case "&Modify"
        Call subEditElement(Val(txtElementNo.Text), Val(txtNodeI.Text), Val(txtNodeJ.Text), cmbSelSec.ListIndex + 1)
        Call subPlotStr
        
    Case "&Delete"
        Call subDeleteElement(Val(txtElementNo.Text))
        Call subPlotStr
End Select

End Sub

Private Sub btnClose_Click()
frmAddElement.Hide
Unload frmAddElement
End Sub


Private Sub txtElementNo_Change()
Dim Value As Long
    If btnAddElement.Caption = "&Add" Then Exit Sub
    Value = Val(txtElementNo.Text)
        cmbSelSec.ListIndex = 0
        Label5.Caption = ""
        txtNodeI.Text = ""
        txtNodeJ.Text = ""
        btnAddElement.Enabled = False
    If Value <= NoOfElements And Value > 0 Then
        cmbSelSec.ListIndex = AsgnSec(Value) - 1
        txtNodeI.Text = Endi(Value)
        txtNodeJ.Text = Endj(Value)
        btnAddElement.Enabled = True
        'txtZcoor.text=Zcoor(value)
    ElseIf txtElementNo.Text <> "" Then
        Label5.Caption = "(Invalid Element Number...)"
    End If
End Sub

Private Sub txtElementNo_GotFocus()
txtElementNo.SelStart = 0
txtElementNo.SelLength = Len(txtElementNo.Text)
End Sub

Private Sub txtNodeI_GotFocus()
txtNodeI.SelStart = 0
txtNodeI.SelLength = Len(txtNodeI.Text)

End Sub


Private Sub txtNodeJ_GotFocus()
txtNodeJ.SelStart = 0
txtNodeJ.SelLength = Len(txtNodeJ.Text)
End Sub
