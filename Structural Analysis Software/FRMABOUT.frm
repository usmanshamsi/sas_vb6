VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About S.A.S"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
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
   ScaleHeight     =   1305
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   5520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label12 
      Caption         =   "Please provide your feedback at usman.shamsi@live.com"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   6480
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   6480
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Please provide your feedback at sasoftware@ymail.com"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   6615
   End
   Begin VB.Label Label10 
      Caption         =   "Contact: +92 333 3582702"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label9 
      Caption         =   "M.Engg. (Structures), NEDUET."
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label8 
      Caption         =   "Muqeet Ahmed"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Contact: +92 334 3506562"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "M.Engg. (Structures), NEDUET."
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Muhammad Usman Shamsi"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "B  U  I  L  D     B Y :"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Release: Alpha"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Version: 1.0"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Structural Analysis Software (S.A.S)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
