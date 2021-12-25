VERSION 5.00
Begin VB.Form frmAddNode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add / Edit Node"
   ClientHeight    =   3105
   ClientLeft      =   7785
   ClientTop       =   5520
   ClientWidth     =   4875
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
   ScaleHeight     =   3105
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   500
      Left            =   3000
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Add"
      Height          =   500
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Coordinates"
      Height          =   2055
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2415
      Begin VB.TextBox txtZCoOr 
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1500
         Width           =   1575
      End
      Begin VB.TextBox txtYCoOr 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   900
         Width           =   1575
      End
      Begin VB.TextBox txtXCoOr 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Z:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Y:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "X:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox txtNodeNo 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Node No:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "(Node do not exist)"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmAddNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit

Private Sub btnAdd_Click()

'CHECK IF THE TEXTBOXES ARE FILLED
'---------------------------------
    If txtXCoOr.Text = "" Then
        MsgBox "Enter Value of X coordinate"
        Exit Sub
    End If
    
    If txtYCoOr.Text = "" Then
        MsgBox "Enter Value of Y coordinate"
        Exit Sub
    End If
    
'ADD OR EDIT NODE ACCORDINGLY
'----------------------------------

Select Case btnAdd.Caption

    Case "&Add"
        Call subAddNode(Val(txtXCoOr.Text), Val(txtYCoOr.Text))
        txtNodeNo.Text = NoofNodes + 1
        subPlotStr
    
    Case "&Modify"
        Call subEditNode(Val(txtNodeNo.Text), Val(txtXCoOr.Text), Val(txtYCoOr.Text))
        subPlotStr
        
    Case "&Delete"
        Call subDeleteNode(Val(txtNodeNo.Text))
        subPlotStr
End Select

End Sub

Private Sub btnClose_Click()
frmAddNode.Hide
Unload frmAddNode
End Sub


Private Sub txtNodeNo_Change()
Dim Value As Long
    If btnAdd.Caption = "&Add" Then Exit Sub

    Value = Val(txtNodeNo.Text)
        Label5.Caption = ""
        txtXCoOr.Text = ""
        txtYCoOr.Text = ""
        btnAdd.Enabled = False
        
    If Value <= NoofNodes And Value > 0 Then
        txtXCoOr.Text = Xcoor(Value)
        txtYCoOr.Text = Ycoor(Value)
        'txtZcoor.text=Zcoor(value)
        btnAdd.Enabled = True
        
    ElseIf txtNodeNo.Text <> "" Then
        
        Label5.Caption = "(Invalid Node Number...)"
    
    End If
    
End Sub

Private Sub txtXCoOr_GotFocus()
txtXCoOr.SelStart = 0
txtXCoOr.SelLength = Len(txtXCoOr.Text)
End Sub



Private Sub txtYCoOr_GotFocus()
txtYCoOr.SelStart = 0
txtYCoOr.SelLength = Len(txtYCoOr.Text)
End Sub

Private Sub txtNodeNo_GotFocus()
txtNodeNo.SelStart = 0
txtNodeNo.SelLength = Len(txtNodeNo.Text)
End Sub
