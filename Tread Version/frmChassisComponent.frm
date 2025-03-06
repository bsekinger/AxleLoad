VERSION 5.00
Begin VB.Form frmChassisComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "What to do with chassis-related components..."
   ClientHeight    =   2652
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6120
   Icon            =   "frmChassisComponent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2652
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   492
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   1452
   End
   Begin VB.OptionButton optImport 
      Caption         =   "Do NOT import the additional components."
      Height          =   372
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   5652
   End
   Begin VB.OptionButton optImport 
      Caption         =   "ADD chassis-related components to existing components"
      Height          =   372
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   5652
   End
   Begin VB.OptionButton optImport 
      Caption         =   "Import components, and REPLACE existing chassis-related components"
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Value           =   -1  'True
      Width           =   5772
   End
   Begin VB.Label lblDescr 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmChassisComponent.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6492
   End
End
Attribute VB_Name = "frmChassisComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sRtn As String

Public Function ImportOption() As String
    
    Me.Show vbModal
    '
    ImportOption = sRtn
    Exit Function

End Function


Private Sub cmdContinue_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    optImport_Click 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'Don't allow closing the form by pressing the 'X'
        Cancel = 1
    End If
End Sub

Private Sub optImport_Click(Index As Integer)
    Select Case Index
    Case 0
        sRtn = "REPLACE"
    Case 1
        sRtn = "ADD"
    Case 2
        sRtn = "IGNORE"
    End Select
End Sub
