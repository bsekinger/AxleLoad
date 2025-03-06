VERSION 5.00
Begin VB.Form frmFillRelationships 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Fill Relationships"
   ClientHeight    =   3996
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6504
   Icon            =   "frmFillRelationships.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3996
   ScaleWidth      =   6504
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   612
      Left            =   5280
      TabIndex        =   12
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   612
      Left            =   3000
      TabIndex        =   11
      Top             =   3240
      Width           =   1092
   End
   Begin VB.TextBox txtEmulMultiplier 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Top             =   2280
      Width           =   852
   End
   Begin VB.TextBox txtANMultiplier 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Top             =   1320
      Width           =   852
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   852
   End
   Begin VB.ComboBox cboComponent 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Text            =   "Component"
      Top             =   360
      Width           =   1932
   End
   Begin VB.Label lblOp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Index           =   2
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   372
   End
   Begin VB.Label lblANDescr 
      Caption         =   "% of Total AN Weight"
      Height          =   276
      Left            =   3960
      TabIndex        =   6
      Top             =   1404
      Width           =   1572
   End
   Begin VB.Label lblOp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   372
   End
   Begin VB.Label lblOp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   492
   End
   Begin VB.Label lblComponent 
      BackStyle       =   0  'Transparent
      Caption         =   "The content weight of..."
      Height          =   276
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label lblOffset 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Amt (kg)"
      Height          =   276
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label lblWeightEmul 
      BackStyle       =   0  'Transparent
      Caption         =   "% of Total Emulsion Weight"
      Height          =   276
      Left            =   3960
      TabIndex        =   9
      Top             =   2400
      Width           =   2172
   End
End
Attribute VB_Name = "frmFillRelationships"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typeFillRelation
    ComponentIndex As Integer
    Offset As Double
    ANMult As Double
    EmulMult As Double
End Type

Private bDirty As Boolean
Private bExitNow As Boolean

Private FR() As typeFillRelation
Private cCurTruck As clsTruck 'used in this form as a link to the active truck

Public Sub Edit(cTruck As clsTruck)
    Set cCurTruck = cTruck
    bExitNow = False
    Me.Show
    If bExitNow Then Unload Me
End Sub


Private Sub cmdApply_Click()
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

Private Sub Form_Load()
    Dim i%
    Dim x%
    Dim intCount As Integer
    Dim intIndex As Integer
    
    lblOffset.Caption = "Minimum Amt (" & cGlobalInfo.MassUnits.Display & ")"
    bDirty = False

    'Count components with fill relationships
    For i% = 1 To cCurTruck.Components.Count
        If cCurTruck.Components(i%).FillRelationShips.Count > 0 Then
            intCount = intCount + 1
        End If
    Next i%
    
    If intCount = 0 Then
        MsgBox "Function not applicable to this truck", vbExclamation
        bExitNow = True
        Exit Sub
    End If


    'Resize holding array
    ReDim FR(1 To intCount) As typeFillRelation
    
    'Fill Combo Box
    For i% = 1 To cCurTruck.Components.Count
        If cCurTruck.Components(i%).FillRelationShips.Count > 0 Then
            cboComponent.AddItem cCurTruck.Components(i%).DisplayName
            intIndex = cboComponent.ListCount - 1
            cboComponent.ItemData(intIndex) = i%
            'Fill holding array
            With FR(intIndex + 1)
                .ComponentIndex = i%
                .Offset = 0
                For x% = 1 To cCurTruck.Components(i%).FillRelationShips.Count
                    If .Offset = 0 Then
                        .Offset = cCurTruck.Components(i%).FillRelationShips(x%).Offset
                    End If
                    If cCurTruck.Components(i%).FillRelationShips(x%).ParentProduct = ptAN Then
                        'AN
                        .ANMult = cCurTruck.Components(i%).FillRelationShips(x%).Multiplier
                    Else
                        'Emulsion
                        .EmulMult = cCurTruck.Components(i%).FillRelationShips(x%).Multiplier
                    End If
                Next x%
            End With
        End If
    Next i%
    
    cboComponent.ListIndex = 0
End Sub

Private Sub cboComponent_Click()
    cboComponent_Change
End Sub

Private Sub cboComponent_KeyPress(KeyAscii As Integer)
    cboComponent_Change
End Sub

Private Sub cboComponent_Change()
    Dim i%
    Dim intIndex As Integer
    Dim CompIndex As Integer

    intIndex = cboComponent.ListIndex + 1
    
    txtANMultiplier.Enabled = False
    txtEmulMultiplier.Enabled = False
    txtOffset.Enabled = False
    txtANMultiplier.Text = "0"
    txtEmulMultiplier.Text = "0"
    
    txtOffset.Text = FR(intIndex).Offset * cGlobalInfo.MassUnits.Multiplier
    txtOffset.Enabled = True
    CompIndex = FR(intIndex).ComponentIndex
    For i% = 1 To cCurTruck.Components(CompIndex).FillRelationShips.Count
        If cCurTruck.Components(CompIndex).FillRelationShips(i%).ParentProduct = ptAN Then
            'AN
            txtANMultiplier.Text = FR(intIndex).ANMult * 100
            txtANMultiplier.Enabled = True
        Else
            'Emulsion
            txtEmulMultiplier.Text = FR(intIndex).EmulMult * 100
            txtEmulMultiplier.Enabled = True
        End If
    Next
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cCurTruck = Nothing
End Sub

Private Sub SaveChanges()
    Dim i%
    Dim x%
    Dim intComp As Integer
    
    'Fill Combo Box
    For i% = 0 To UBound(FR) - 1
        'Fill holding array
        With FR(i% + 1)
            intComp = .ComponentIndex
            For x% = 1 To cCurTruck.Components(intComp).FillRelationShips.Count
                If cCurTruck.Components(intComp).FillRelationShips(x%).ParentProduct = ptAN Then
                    'AN
                    cCurTruck.Components(intComp).FillRelationShips(x%).Multiplier = .ANMult
                Else
                    'Emulsion
                    cCurTruck.Components(intComp).FillRelationShips(x%).Multiplier = .EmulMult
                End If
                cCurTruck.Components(intComp).FillRelationShips(x%).Offset = .Offset
                .Offset = 0 'only one offset is allowable
            Next x%
        End With
    Next i%
    bDirty = False 'holding variables moved to truck class
    bTruckFileDirty = True 'Truck file now dirty
End Sub



Private Sub txtOffset_Change()
    If Not txtOffset.Enabled Then Exit Sub
    FR(cboComponent.ListIndex + 1).Offset = Var2Dbl(txtOffset.Text) / cGlobalInfo.MassUnits.Multiplier
End Sub

Private Sub txtANMultiplier_Change()
    If Not txtANMultiplier.Enabled Then Exit Sub
    
    FR(cboComponent.ListIndex + 1).ANMult = Var2Dbl(txtANMultiplier.Text) / 100
End Sub

Private Sub txtEmulMultiplier_Change()
    If Not txtEmulMultiplier.Enabled Then Exit Sub
    
    FR(cboComponent.ListIndex + 1).EmulMult = Var2Dbl(txtEmulMultiplier.Text) / 100
End Sub

Private Sub txtANMultiplier_KeyPress(KeyAscii As Integer)
    bDirty = True
End Sub

Private Sub txtEmulMultiplier_KeyPress(KeyAscii As Integer)
    bDirty = True
End Sub

Private Sub txtOffset_KeyPress(KeyAscii As Integer)
    bDirty = True
End Sub
