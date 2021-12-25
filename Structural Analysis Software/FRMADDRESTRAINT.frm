VERSION 5.00
Begin VB.Form frmAddRestraint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apply Restraint"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
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
   ScaleHeight     =   4620
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptCustomSupport 
      Caption         =   "Custom"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton OptRoller 
      Caption         =   "Roller"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton OptFixed 
      Caption         =   "Fixed"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton optPinned 
      Caption         =   "Pinned"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   500
      Left            =   2040
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton btnAppRes 
      Caption         =   "&Apply"
      Height          =   500
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame FrameRotation 
      Caption         =   "Rotation"
      Height          =   2055
      Left            =   2040
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
      Begin VB.CheckBox chkRz 
         Caption         =   "About Z"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkRy 
         Caption         =   "About Y"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkRx 
         Caption         =   "About X"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame FrameTranslation 
      Caption         =   "Translation"
      Height          =   2055
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
      Begin VB.CheckBox chkTz 
         Caption         =   "Along Z"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkTy 
         Caption         =   "Along Y"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkTx 
         Caption         =   "Along X"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txtNodeNo 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Node No:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddRestraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit

Private Sub btnAppRes_Click()
    Call subApplyRestraint(Val(txtNodeNo.Text), _
                            chkTx.Value, chkTy.Value, _
                            chkRz.Value)
End Sub

Private Sub btnClose_Click()
    frmAddRestraint.Hide
    Unload frmAddRestraint
End Sub

Private Sub OptCustomSupport_Click()
    FrameTranslation.Enabled = OptCustomSupport.Value
End Sub

Private Sub OptFixed_Click()
    chkTx.Value = 1
    chkTy.Value = 1
    chkTz.Value = 1
    chkRx.Value = 1
    chkRy.Value = 1
    chkRz.Value = 1
End Sub

Private Sub optPinned_Click()
    chkTx.Value = 1
    chkTy.Value = 1
    chkTz.Value = 1
    chkRx.Value = 0
    chkRy.Value = 0
    chkRz.Value = 0
End Sub

Private Sub OptRoller_Click()
    chkTx.Value = 0
    chkTy.Value = 1
    chkTz.Value = 0
    chkRx.Value = 0
    chkRy.Value = 0
    chkRz.Value = 0
End Sub

Private Sub txtNodeNo_Change()
Dim Value As Long
    
    Value = Val(txtNodeNo.Text)
    Label5.Caption = ""
    If Value > NoofNodes Then
        Label5.Caption = "(Invalid Node Number...)"
        Exit Sub
    End If
    
    OptFixed.Value = 0
    optPinned.Value = 0
    OptRoller.Value = 0
    OptCustomSupport.Value = 0
    chkTx.Value = 0
    chkTy.Value = 0
    chkTz.Value = 0
    chkRx.Value = 0
    chkRy.Value = 0
    chkRz.Value = 0
    
    If Value <= NoofNodes And Value > 0 Then
    
        OptFixed.Value = 0
        optPinned.Value = 0
        OptRoller.Value = 0
        OptCustomSupport.Value = 0
        
        chkTx.Value = Val(TxRest(Value))
        chkTy.Value = Val(TyRest(Value))
        'chkTz.Value = 0
        'chkRx.Value = 0
        'chkRy.Value = 0
        chkRz.Value = Val(RzRest(Value))

        
    End If
    

End Sub

Private Sub txtNodeNo_GotFocus()
    txtNodeNo.SelStart = 0
    txtNodeNo.SelLength = Len(txtNodeNo.Text)
End Sub
