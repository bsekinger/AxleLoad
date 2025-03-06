VERSION 5.00
Begin VB.Form frmEditCoefficients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Coefficients"
   ClientHeight    =   3612
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   2064
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3612
   ScaleWidth      =   2064
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "OK"
      Height          =   372
      Left            =   480
      TabIndex        =   12
      Top             =   3120
      Width           =   1452
   End
   Begin VB.TextBox txtK 
      Height          =   315
      Index           =   5
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   1452
   End
   Begin VB.TextBox txtK 
      Height          =   315
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   720
      Width           =   1452
   End
   Begin VB.TextBox txtK 
      Height          =   315
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   1452
   End
   Begin VB.TextBox txtK 
      Height          =   315
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   1452
   End
   Begin VB.TextBox txtK 
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1452
   End
   Begin VB.TextBox txtK 
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   1452
   End
   Begin VB.Label lblK5 
      Caption         =   "K5"
      Height          =   276
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   372
   End
   Begin VB.Label lblK4 
      Caption         =   "K4"
      Height          =   276
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   372
   End
   Begin VB.Label lblK3 
      Caption         =   "K3"
      Height          =   276
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   372
   End
   Begin VB.Label lblK2 
      Caption         =   "K2"
      Height          =   276
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   372
   End
   Begin VB.Label lblK1 
      Caption         =   "K1"
      Height          =   276
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   372
   End
   Begin VB.Label lblK 
      Caption         =   "K0"
      Height          =   276
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   372
   End
End
Attribute VB_Name = "frmEditCoefficients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dblK(5) As Double
Private bOK As Boolean

Public Function EditCoefficients(ByRef colInput As Collection, frmParent As Form) As Collection
    On Error GoTo errHandler
    'Called from elsewhere to edit either StickLength or ContentCG coefficients
    Dim i%
    Dim colK As Collection
    
    bOK = False
    
    'Make sure we have 6 valid points to deal with
    For i% = colInput.Count To 5
        colInput.Add 0, CStr(i%)
    Next
    'Assign input collection to local array
    For i% = 0 To 5
        dblK(i%) = colInput(CStr(i%))
    Next
    'Now show the form (modal)
    Me.Show vbModal, frmParent
    'Form was closed...
    If bOK Then
        'User pressed 'OK' so replace colCoeff with contents of form
        Set colK = New Collection
        For i% = 0 To 5
            colK.Add dblK(i%), CStr(i%)
        Next
        Set EditCoefficients = colK
    Else
        'User clicked on 'close window', no changes will be returned
        Set EditCoefficients = colInput
    End If
    Exit Function
errHandler:
    ErrorIn "frmEditCoefficients.EditCoefficients(colInput,frmParent)", Array(colInput, frmParent)
End Function


Private Sub cmdExit_Click()
    bOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i%
    
    'Pre-fill with zeros
    For i% = 0 To 5
        txtK(i%).Text = dblK(i%)
    Next i%
End Sub


Private Sub txtK_Validate(index As Integer, Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtK(index).Text) Then
        sMsg = "Please enter a valid number"
        MsgBox sMsg, vbExclamation, "Invalid Input"
        Cancel = True
    Else
        'Save value to collection
        dblK(index) = txtK(index).Text
    End If
End Sub
