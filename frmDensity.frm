VERSION 5.00
Begin VB.Form frmDensity 
   Caption         =   "Change Density Values"
   ClientHeight    =   4200
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4068
   Icon            =   "frmDensity.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4068
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   852
      Left            =   2640
      TabIndex        =   17
      Top             =   3240
      Width           =   1332
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   852
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1332
   End
   Begin VB.TextBox txtDensity 
      Height          =   288
      Index           =   7
      Left            =   1920
      TabIndex        =   10
      Top             =   2760
      Width           =   1092
   End
   Begin VB.TextBox txtDensity 
      Height          =   288
      Index           =   6
      Left            =   1920
      TabIndex        =   9
      Top             =   2400
      Width           =   1092
   End
   Begin VB.TextBox txtDensity 
      Height          =   288
      Index           =   5
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   1092
   End
   Begin VB.TextBox txtDensity 
      Height          =   288
      Index           =   4
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   1092
   End
   Begin VB.TextBox txtDensity 
      Height          =   288
      Index           =   3
      Left            =   1920
      TabIndex        =   6
      Top             =   1320
      Width           =   1092
   End
   Begin VB.TextBox txtDensity 
      Height          =   288
      Index           =   2
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   1092
   End
   Begin VB.TextBox txtDensity 
      Height          =   288
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp_7"
      Height          =   276
      Index           =   7
      Left            =   600
      TabIndex        =   15
      Top             =   2760
      Width           =   1212
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp_6"
      Height          =   276
      Index           =   6
      Left            =   600
      TabIndex        =   14
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp_5"
      Height          =   276
      Index           =   5
      Left            =   600
      TabIndex        =   13
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp_4"
      Height          =   276
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp_3"
      Height          =   276
      Index           =   3
      Left            =   600
      TabIndex        =   11
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp_2"
      Height          =   276
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp_1"
      Height          =   276
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   1212
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   600
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblDensity 
      BackStyle       =   0  'Transparent
      Caption         =   "Density (aka SG)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label lblComponent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Component"
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
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
End
Attribute VB_Name = "frmDensity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDirty As Boolean

Private cCurTruck As clsTruck 'used in this form as a link to the active truck

Public Sub Edit(cTruck As clsTruck)
    Set cCurTruck = cTruck
    Me.Show
End Sub


Private Sub Form_Load()
    Dim i%
    Dim intNum
    Dim bHasAN As Boolean
    Dim bHasEmul As Boolean
    Dim cType As ComponentContentsType
    
    For i% = 1 To 7
        lblComponent(i%).Visible = False
        txtDensity(i%).Visible = False
    Next i%
    
    intNum = 0
    For i% = 1 To cCurTruck.Body.Tanks.Count
        If cCurTruck.Body.Tanks(i%).CurTankUse = ttAN Then
            bHasAN = True
        Else
            bHasEmul = True
        End If
    Next i%
    
    If bHasAN Then
        intNum = intNum + 1
        lblComponent(intNum).Caption = "AN Prill"
        txtDensity(intNum).Tag = "AN"
        txtDensity(intNum).Text = cGlobalInfo.DensityAN
        lblComponent(intNum).Visible = True
        txtDensity(intNum).Visible = True
    End If
    
    If bHasEmul Then
        intNum = intNum + 1
        lblComponent(intNum).Caption = "Emulsion"
        txtDensity(intNum).Tag = "Emulsion"
        txtDensity(intNum).Text = cGlobalInfo.DensityEmul
        lblComponent(intNum).Visible = True
        txtDensity(intNum).Visible = True
    End If
    
    For i% = 1 To cCurTruck.Components.Count
        Select Case cCurTruck.Components(i%).ContentsType
        Case ctFuel
            intNum = intNum + 1
            lblComponent(intNum).Caption = "Fuel"
            txtDensity(intNum).Tag = "Fuel"
            txtDensity(intNum).Text = cGlobalInfo.DensityFuel
            lblComponent(intNum).Visible = True
            txtDensity(intNum).Visible = True
        Case ctWater
            intNum = intNum + 1
            lblComponent(intNum).Caption = "Water"
            txtDensity(intNum).Tag = "Water"
            txtDensity(intNum).Text = cGlobalInfo.DensityFuel
            lblComponent(intNum).Visible = True
            txtDensity(intNum).Visible = True
        Case ctGasA
            intNum = intNum + 1
            lblComponent(intNum).Caption = "Gas A"
            txtDensity(intNum).Tag = "Gas A"
            txtDensity(intNum).Text = cGlobalInfo.DensityGasA
            lblComponent(intNum).Visible = True
            txtDensity(intNum).Visible = True
        Case ctGasB
            intNum = intNum + 1
            lblComponent(intNum).Caption = "Gas B"
            txtDensity(intNum).Tag = "Gas B"
            txtDensity(intNum).Text = cGlobalInfo.DensityGasB
            lblComponent(intNum).Visible = True
            txtDensity(intNum).Visible = True
        Case ctAdditive
            intNum = intNum + 1
            lblComponent(intNum).Caption = "Additive"
            txtDensity(intNum).Tag = "Additive"
            txtDensity(intNum).Text = cGlobalInfo.DensityAdditive
            lblComponent(intNum).Visible = True
            txtDensity(intNum).Visible = True
        End Select
    Next i%
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cCurTruck = Nothing
End Sub


Private Sub cmdApply_Click()
    Dim a$
    Dim sFile As String
    
    'Save changes to current session
    SaveChanges
    a$ = "Changes have been saved for this session." & vbCrLf & vbCrLf & _
         "Save these changes for future sessions also?"
    If MsgBox(a$, vbYesNo) = vbYes Then
        'Save GlobalInfo
        sFile = AddBackslash(App.Path) & "GlobalInfo.xml"
        SaveGlobalInfo sFile
    End If
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If bDirty Then
        If MsgBox("Save Changes?", vbYesNo) = vbYes Then
            SaveChanges
        End If
    End If
    Unload Me
End Sub

Private Sub SaveChanges()
    'Saves entered data to 'cGlobalInfo' (i.e. current session)
    Dim i%
    
    For i% = 1 To 7
        If txtDensity(i%).Visible = True Then
            Select Case txtDensity(i%).Tag
            Case "AN"
                cGlobalInfo.DensityAN = Var2Dbl(txtDensity(i%).Text)
            Case "Emulsion"
                cGlobalInfo.DensityEmul = Var2Dbl(txtDensity(i%).Text)
            Case "Fuel"
                cGlobalInfo.DensityFuel = Var2Dbl(txtDensity(i%).Text)
            Case "Water"
                cGlobalInfo.DensityFuel = Var2Dbl(txtDensity(i%).Text)
            Case "Gas A"
                cGlobalInfo.DensityGasA = Var2Dbl(txtDensity(i%).Text)
            Case "Gas B"
                cGlobalInfo.DensityGasB = Var2Dbl(txtDensity(i%).Text)
            Case "Additive"
                cGlobalInfo.DensityAdditive = Var2Dbl(txtDensity(i%).Text)
            End Select
        End If
    Next i%
End Sub
