VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About..."
   ClientHeight    =   2916
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   3468
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2916
   ScaleWidth      =   3468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   1248
      TabIndex        =   4
      Top             =   2280
      Width           =   972
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      Height          =   1212
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3252
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "x.x.x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   492
   End
   Begin VB.Label lblProgram 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Program version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1692
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Client Version"
      Height          =   276
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3252
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unibody AxleLoad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3252
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = App.LegalCopyright
    
    #If TREADVERSION = 1 Then ' Only compiled in Tread Version of the program -----------------
        lblType.Caption = "Tread Version"
    #Else
        lblType.Caption = "Client Version"
    #End If
End Sub

