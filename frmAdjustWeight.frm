VERSION 5.00
Begin VB.Form frmAdjustWeight 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjust Empty Weight"
   ClientHeight    =   3972
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6588
   Icon            =   "frmAdjustWeight.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3972
   ScaleWidth      =   6588
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   852
      Left            =   2640
      TabIndex        =   8
      Top             =   3000
      Width           =   1332
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   852
      Left            =   5160
      TabIndex        =   7
      Top             =   3000
      Width           =   1332
   End
   Begin VB.TextBox txtRear 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   3540
      Width           =   1332
   End
   Begin VB.TextBox txtFront 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   3036
      Width           =   1332
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Adder Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   1332
   End
   Begin VB.Label lblMassUnits 
      BackStyle       =   0  'Transparent
      Caption         =   "kg"
      Height          =   276
      Left            =   2040
      TabIndex        =   5
      Top             =   2760
      Width           =   492
   End
   Begin VB.Label lblRear 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rear"
      Height          =   276
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   612
   End
   Begin VB.Label lblFront 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Front"
      Height          =   276
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   612
   End
   Begin VB.Label lblWarning 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAdjustWeight.frx":08CA
      ForeColor       =   &H80000008&
      Height          =   2292
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6372
   End
End
Attribute VB_Name = "frmAdjustWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cCurTruck As clsTruck 'used in this form as a link to the active truck

Public Sub Edit(cTruck As clsTruck)
    Set cCurTruck = cTruck
    Me.Show
End Sub

Private Sub Form_Load()
    Dim dblVal As Double
    
    lblMassUnits.Caption = "(" & cGlobalInfo.MassUnits.Display & ")"
    dblVal = cCurTruck.WtAdjustFront * cGlobalInfo.MassUnits.Multiplier
    txtFront.Text = Format(dblVal, "#,##0")
    
    dblVal = cCurTruck.WtAdjustRear * cGlobalInfo.MassUnits.Multiplier
    txtRear.Text = Format(dblVal, "#,##0")
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdApply_Click()
    cCurTruck.WtAdjustFront = Var2Dbl(txtFront.Text) / cGlobalInfo.MassUnits.Multiplier
    cCurTruck.WtAdjustRear = Var2Dbl(txtRear.Text) / cGlobalInfo.MassUnits.Multiplier
    bTruckFileDirty = True
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cCurTruck = Nothing
End Sub
