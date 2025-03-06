VERSION 5.00
Begin VB.Form frmAutoTankFilling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dependent Tank Filling"
   ClientHeight    =   1596
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   2808
   Icon            =   "frmAutoTankFilling.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1596
   ScaleWidth      =   2808
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboFillType 
      BackColor       =   &H00C0FFC0&
      Height          =   288
      Index           =   1
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label lblComponent 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label lblDensity 
      BackStyle       =   0  'Transparent
      Caption         =   "Set level..."
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
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1092
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   120
      X2              =   2640
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblContentsType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ContentsType_1"
      Height          =   276
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1212
   End
End
Attribute VB_Name = "frmAutoTankFilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDirty As Boolean
Private intNumDependent As Integer

Private cCurTruck As clsTruck 'used in this form as a link to the active truck



Public Sub Options(cTruck As clsTruck)
    Set cCurTruck = cTruck
    intNumDependent = 0
    Me.Show vbModal, frmMain
    
    Set cCurTruck = Nothing
End Sub



Private Sub Form_Load()
    Dim i%, x%
    Dim CompIndex As Integer
    Dim bTemp As Boolean
    Dim cType As ComponentContentsType
    
    'Create Controls & resize form
    For i% = 1 To cCurTruck.Components.Count
        If cCurTruck.Components(i%).ContentsType <> ctOther _
         And cCurTruck.Components(i%).ContentsType <> ctNone Then
            'This is a dependent-style tank
            intNumDependent = intNumDependent + 1
            bTemp = False
            If intNumDependent > 1 Then
                'see if contents-type is already represented
                If intNumDependent > 1 Then
                    cType = cCurTruck.Components(i%).ContentsType
                    For x% = 1 To intNumDependent - 1
                        If cboFillType(x%).Tag = cType Then
                            'An earlier-displayed tank already shows this contents-type
                            intNumDependent = intNumDependent - 1 'back up the counter
                            bTemp = True
                            Exit For
                        End If
                    Next
                End If
                
                'only add combo box if contents-type is not already represented
                If Not bTemp Then
                    Load lblContentsType(intNumDependent)
                    With lblContentsType(intNumDependent)
                        .Left = lblContentsType(intNumDependent - 1).Left
                        .Top = lblContentsType(intNumDependent - 1).Top + 360
                        .Visible = True
                    End With
                    Load cboFillType(intNumDependent)
                    With cboFillType(intNumDependent)
                        .Left = cboFillType(intNumDependent - 1).Left
                        .Top = cboFillType(intNumDependent - 1).Top + 360
                        .Visible = True
                    End With
                End If
            End If
            'set default captions, lists
            If Not bTemp Then 'if contents-type is not already represented
                'Load combo-box w/ choices
                cboFillType(intNumDependent).AddItem "Auto", 0
                cboFillType(intNumDependent).AddItem "Empty", 1
                cboFillType(intNumDependent).AddItem "Full", 2
                cboFillType(intNumDependent).Tag = cCurTruck.Components(i%).ContentsType 'tag = component index
                'Label
                lblContentsType(intNumDependent).Caption = cCurTruck.Components(i%).ContentsTypeString
            End If
        End If
    Next i%
    cmdApply.Top = lblContentsType(intNumDependent).Top + 480
    cmdCancel.Top = cmdApply.Top
    Me.Height = 2016 + (cmdApply.Top - 960)


    'Set existing choices
    For i% = 1 To intNumDependent
        Select Case CInt(cboFillType(i%).Tag)
        Case ctFuel
            cboFillType(i%).ListIndex = cGlobalInfo.FillMethod_Fuel
        Case ctWater
            cboFillType(i%).ListIndex = cGlobalInfo.FillMethod_Water
        Case ctGasA
            cboFillType(i%).ListIndex = cGlobalInfo.FillMethod_GasA
        Case ctGasB
            cboFillType(i%).ListIndex = cGlobalInfo.FillMethod_GasB
        Case ctAdditive
            cboFillType(i%).ListIndex = cGlobalInfo.FillMethod_Additive
        End Select
    Next i%
    bDirty = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        cmdCancel_Click
    End If
End Sub

Private Sub cboFillType_Click(Index As Integer)
     cboFillType_Change Index
End Sub


Private Sub cboFillType_Change(Index As Integer)
    bDirty = True
    With cboFillType(Index)
        Select Case .ListIndex
        Case 0 '0=Auto
            .BackColor = vbWindowBackground
        Case 1 '1=Empty
            .BackColor = &HC0C0FF    'Light Red
        Case 2 '2=Full
            .BackColor = &HC0FFC0    'Light Green
        End Select
    End With
End Sub


Private Sub cmdApply_Click()
    Dim a$
    Dim sFile As String
    
    'Save changes to current session
    SaveChanges
    
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
    
    For i% = 1 To intNumDependent
        Select Case CInt(cboFillType(i%).Tag)
        Case ctFuel
            cGlobalInfo.FillMethod_Fuel = cboFillType(i%).ListIndex
        Case ctWater
            cGlobalInfo.FillMethod_Water = cboFillType(i%).ListIndex
        Case ctGasA
            cGlobalInfo.FillMethod_GasA = cboFillType(i%).ListIndex
        Case ctGasB
            cGlobalInfo.FillMethod_GasB = cboFillType(i%).ListIndex
        Case ctAdditive
            cGlobalInfo.FillMethod_Additive = cboFillType(i%).ListIndex
        End Select
    Next i%
End Sub


