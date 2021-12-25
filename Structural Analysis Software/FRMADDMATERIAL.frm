VERSION 5.00
Begin VB.Form frmAddMaterial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Material"
   ClientHeight    =   2565
   ClientLeft      =   7245
   ClientTop       =   3045
   ClientWidth     =   7470
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
   ScaleHeight     =   2565
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMaterialName 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtMatCoeffTher 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtMatShearMod 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtMatElasMod 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton btnAddMat 
      Caption         =   "&Add"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtMatID 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Material Name:"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   300
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Material ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Elastic Modulus (&E):"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Shear Modulus (&G):"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1500
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Coeff. of  Thermal Expansion:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2100
      Width           =   3015
   End
End
Attribute VB_Name = "frmAddMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option base 1
Option Explicit

Private Sub btnAddMat_Click()

    'CHECK IF THE INPUT TEXTBOXES ARE FILLED
        If txtMaterialName.Text = "" Then
            MsgBox "Enter Material Name"
            Exit Sub
        End If
        
        
        If txtMatElasMod.Text = "" Then
            MsgBox "Enter Elastic Modulus"
            Exit Sub
        End If
        
        If txtMatShearMod.Text = "" Then
            MsgBox "Enter Shear Modulus"
            Exit Sub
        End If
        
        If txtMatCoeffTher.Text = "" Then
            MsgBox "Enter Coeff. of Thermal"
            Exit Sub
        End If

    'ADD OF EDIT MATERIAL ACCORDINGLY
    Select Case btnAddMat.Caption
    
        Case "&Add"
            Call subAddMaterial(txtMaterialName.Text, _
                                Val(txtMatElasMod.Text), _
                                Val(txtMatShearMod.Text), _
                                Val(txtMatCoeffTher.Text))
            txtMatID.Text = NoOfMaterials + 1
    
        Case "&Modify"
            Call subEditMaterial(Val(txtMatID.Text), _
                                txtMaterialName.Text, _
                                Val(txtMatElasMod.Text), _
                                Val(txtMatShearMod.Text), _
                                Val(txtMatCoeffTher.Text))

            
    End Select
End Sub


Private Sub btnClose_Click()
frmAddMaterial.Hide
Unload frmAddMaterial
End Sub


Private Sub txtMatCoeffTher_GotFocus()
txtMatCoeffTher.SelStart = 0
txtMatCoeffTher.SelLength = Len(txtMatCoeffTher.Text)
End Sub

Private Sub txtMatElasMod_GotFocus()
txtMatElasMod.SelStart = 0
txtMatElasMod.SelLength = Len(txtMatElasMod.Text)
End Sub

Private Sub txtMaterialName_GotFocus()
txtMaterialName.SelStart = 0
txtMaterialName.SelLength = Len(txtMaterialName.Text)
End Sub

Private Sub txtMatID_Change()
    If btnAddMat.Caption = "&Add" Then Exit Sub
Dim Value As Long
    Value = Val(txtMatID.Text)
        txtMaterialName.Text = ""
        txtMatElasMod.Text = ""
        txtMatShearMod.Text = ""
        txtMatCoeffTher.Text = ""
        btnAddMat.Enabled = False
        
    If Value <= NoOfMaterials And Value > 0 Then
        txtMaterialName.Text = MaterialName(Value)
        txtMatElasMod.Text = MatElasMod(Value)
        txtMatShearMod.Text = MatShearMod(Value)
        txtMatCoeffTher.Text = MatCoeffTher(Value)
        btnAddMat.Enabled = True
        
    End If
End Sub

Private Sub txtMatID_GotFocus()
txtMatID.SelStart = 0
txtMatID.SelLength = Len(txtMatID.Text)
End Sub

Private Sub txtMatShearMod_GotFocus()
txtMatShearMod.SelStart = 0
txtMatShearMod.SelLength = Len(txtMatShearMod.Text)
End Sub
