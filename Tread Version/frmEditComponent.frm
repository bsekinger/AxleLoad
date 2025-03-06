VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEditComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Component"
   ClientHeight    =   7668
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7668
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLocation 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Left            =   240
      TabIndex        =   59
      Top             =   1560
      Width           =   2292
      Begin VB.TextBox txtCompOffset 
         Height          =   315
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "Enter the distance from the ""Locating Reference point"" to the Component's origin.  "
         Top             =   1236
         Width           =   1332
      End
      Begin VB.ComboBox cboLocationReference 
         Height          =   288
         Left            =   120
         TabIndex        =   60
         Text            =   "Combo1"
         ToolTipText     =   "'Chassis' reference is the front-most axle."
         Top             =   516
         Width           =   1812
      End
      Begin VB.Label lblDistanceUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "DistUnits"
         Height          =   276
         Index           =   0
         Left            =   1560
         TabIndex        =   64
         Top             =   1356
         Width           =   732
      End
      Begin VB.Label lblLocationReference 
         Caption         =   "Reference Point"
         Height          =   276
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1812
      End
      Begin VB.Label lblOffset 
         Caption         =   "Offset from Reference Point"
         Height          =   276
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   960
         Width           =   2052
      End
   End
   Begin VB.Frame fraLayout 
      Caption         =   "Layout Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3012
      Left            =   2760
      TabIndex        =   50
      Top             =   2280
      Width           =   2892
      Begin VB.TextBox txtCurbSideStd 
         Height          =   315
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "Curb view of chassis or a std-mt body OR street view of a rev-mt body"
         Top             =   2436
         Width           =   2652
      End
      Begin VB.TextBox txtStreetSideStd 
         Height          =   315
         Left            =   120
         TabIndex        =   55
         ToolTipText     =   "Street view of chassis or a std-mt body OR curb view of a rev-mt body"
         Top             =   1596
         Width           =   2652
      End
      Begin VB.ComboBox cboPlacementAllowable 
         Height          =   288
         Left            =   100
         TabIndex        =   53
         Text            =   "Center"
         Top             =   840
         Width           =   1300
      End
      Begin VB.ComboBox cboPlacement 
         Height          =   288
         Left            =   1500
         TabIndex        =   51
         Text            =   "Not Placed"
         Top             =   840
         Width           =   1300
      End
      Begin VB.Label lblCurbSideStd 
         BackStyle       =   0  'Transparent
         Caption         =   "CurbSideStd Drawing"
         Height          =   276
         Left            =   120
         TabIndex        =   58
         Top             =   2160
         Width           =   2052
      End
      Begin VB.Label lblStreetSideStd 
         BackStyle       =   0  'Transparent
         Caption         =   "StreetSideStd Drawing"
         Height          =   276
         Left            =   120
         TabIndex        =   56
         Top             =   1320
         Width           =   1812
      End
      Begin VB.Label lblPlacementAllowable 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditComponent.frx":0000
         Height          =   372
         Left            =   240
         TabIndex        =   54
         Top             =   410
         Width           =   852
      End
      Begin VB.Label lblPlacement 
         BackStyle       =   0  'Transparent
         Caption         =   "Placement"
         Height          =   276
         Left            =   1560
         TabIndex        =   52
         Top             =   600
         Width           =   972
      End
   End
   Begin VB.TextBox txtNotes 
      Height          =   1152
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   48
      Top             =   5640
      Width           =   5292
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save && Exit"
      Height          =   612
      Left            =   9480
      TabIndex        =   46
      Top             =   6960
      Width           =   1572
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   $"frmEditComponent.frx":0018
      Height          =   612
      Left            =   5880
      TabIndex        =   45
      Top             =   6960
      Width           =   1572
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   612
      Left            =   120
      TabIndex        =   44
      Top             =   6960
      Width           =   1572
   End
   Begin VB.TextBox txtFullName 
      Height          =   315
      Left            =   240
      TabIndex        =   41
      Top             =   396
      Width           =   6732
   End
   Begin VB.TextBox txtDisplayName 
      Height          =   315
      Left            =   240
      TabIndex        =   40
      Top             =   1080
      Width           =   2532
   End
   Begin VB.Frame fraEmpty 
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   240
      TabIndex        =   33
      Top             =   3480
      Width           =   2292
      Begin VB.TextBox txtEmptyWeight 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   636
         Width           =   1212
      End
      Begin VB.TextBox txtEmptyCG 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Enter the distance from the Component's origin to the Component's CG."
         Top             =   1356
         Width           =   1212
      End
      Begin VB.Label lblLocNote 
         BackStyle       =   0  'Transparent
         Caption         =   "(Rel. to Component Origin)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   720
         TabIndex        =   49
         Top             =   1100
         Width           =   1452
      End
      Begin VB.Label lblEmptyWeight 
         Caption         =   "Weight"
         Height          =   276
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label lblEmptyCG 
         Caption         =   "CG Loc."
         Height          =   276
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   612
      End
      Begin VB.Label lblDistanceUnits 
         Caption         =   "DistUnits"
         Height          =   276
         Index           =   1
         Left            =   1440
         TabIndex        =   37
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label lblMassUnits 
         Caption         =   "Mass"
         Height          =   276
         Index           =   0
         Left            =   1440
         TabIndex        =   36
         Top             =   720
         Width           =   732
      End
   End
   Begin VB.Frame fraContents 
      Caption         =   "Contents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5892
      Left            =   5880
      TabIndex        =   0
      Top             =   960
      Width           =   5172
      Begin VB.ComboBox cboContentsType 
         Height          =   288
         Left            =   120
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   636
         Width           =   1812
      End
      Begin VB.TextBox txtDefaultWtContents 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Only used if ContentsType=Other. Blank = don't adjust by weight"
         Top             =   1476
         Width           =   1212
      End
      Begin VB.TextBox txtDefaultVolContents 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Only used if ContentsType=Other. Blank = don't adjust by volume"
         Top             =   2316
         Width           =   1212
      End
      Begin VB.TextBox txtDensityContents 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Only used if ContentsType=Other"
         Top             =   3156
         Width           =   1212
      End
      Begin VB.CheckBox chkUsesSightGauge 
         Caption         =   "Uses Sight Gauge"
         Height          =   252
         Left            =   2520
         TabIndex        =   18
         ToolTipText     =   "Check to display fractional fill instead of StickLength"
         Top             =   2160
         Width           =   2172
      End
      Begin VB.TextBox txtVolume 
         Height          =   315
         Left            =   2520
         TabIndex        =   17
         ToolTipText     =   "Not Used if ContentsType=None"
         Top             =   636
         Width           =   1452
      End
      Begin VB.CommandButton cmdStickLength 
         Caption         =   "Edit Coefficients"
         Height          =   372
         Left            =   2520
         TabIndex        =   16
         Top             =   1440
         Width           =   1572
      End
      Begin VB.CommandButton cmdContentCG 
         Caption         =   "Edit Coefficients"
         Height          =   372
         Left            =   2520
         TabIndex        =   15
         Top             =   3120
         Width           =   1692
      End
      Begin VB.Frame fraAN 
         Height          =   2052
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   2412
         Begin VB.TextBox txtOffset 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   1596
            Width           =   1332
         End
         Begin VB.TextBox txtMultiplier 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   876
            Width           =   1332
         End
         Begin VB.CheckBox chkAN 
            Caption         =   "AN Relationship"
            Height          =   372
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1932
         End
         Begin VB.Label lblMassUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass"
            Height          =   276
            Index           =   2
            Left            =   1560
            TabIndex        =   14
            Top             =   1680
            Width           =   732
         End
         Begin VB.Label lblOffset 
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
            Height          =   276
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1332
         End
         Begin VB.Label lblMultiplier 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier"
            Height          =   276
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1332
         End
      End
      Begin VB.Frame fraEmulsion 
         Height          =   2052
         Left            =   2640
         TabIndex        =   1
         Top             =   3720
         Width           =   2412
         Begin VB.TextBox txtOffset 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   1596
            Width           =   1332
         End
         Begin VB.TextBox txtMultiplier 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   876
            Width           =   1332
         End
         Begin VB.CheckBox chkEmulsion 
            Caption         =   "Emulsion Relationship"
            Height          =   372
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   1932
         End
         Begin VB.Label lblMassUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass"
            Height          =   280
            Index           =   3
            Left            =   1560
            TabIndex        =   7
            Top             =   1680
            Width           =   732
         End
         Begin VB.Label lblOffset 
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
            Height          =   280
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1332
         End
         Begin VB.Label lblMultiplier 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier"
            Height          =   280
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1332
         End
      End
      Begin VB.Label lblContentsType 
         Caption         =   "Type"
         Height          =   276
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   2172
      End
      Begin VB.Label lblDefaultWtContents 
         Caption         =   "Default Weight"
         Height          =   276
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label lblMassUnits 
         Caption         =   "Mass"
         Height          =   276
         Index           =   1
         Left            =   1440
         TabIndex        =   30
         Top             =   1596
         Width           =   732
      End
      Begin VB.Label lblDefaultVolContents 
         Caption         =   "Default Volume"
         Height          =   276
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   1212
      End
      Begin VB.Label lblVolumeUnits 
         Caption         =   "VolUnits"
         Height          =   276
         Index           =   0
         Left            =   1440
         TabIndex        =   28
         Top             =   2436
         Width           =   732
      End
      Begin VB.Label lblDensityContents 
         Caption         =   "Specific Gravity"
         Height          =   276
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "*same thing as kg/liter"
         Top             =   2880
         Width           =   1212
      End
      Begin VB.Label lblVolume 
         Caption         =   "Capacity (Volume)"
         Height          =   276
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label lblVolumeUnits 
         Caption         =   "VolUnits"
         Height          =   276
         Index           =   1
         Left            =   4080
         TabIndex        =   25
         Top             =   756
         Width           =   732
      End
      Begin VB.Label lblStickLength 
         Caption         =   "Stick Length Equation"
         Height          =   276
         Left            =   2520
         TabIndex        =   24
         Top             =   1200
         Width           =   1692
      End
      Begin VB.Label lblContentCG 
         Caption         =   "Content CG Equation"
         Height          =   276
         Left            =   2520
         TabIndex        =   23
         Top             =   2880
         Width           =   1812
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   8520
      Top             =   120
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Label lblDisplayName 
      Caption         =   "Display Name"
      Height          =   276
      Left            =   240
      TabIndex        =   43
      Top             =   840
      Width           =   1212
   End
   Begin VB.Label lblFullName 
      Caption         =   "Full Name"
      Height          =   276
      Left            =   240
      TabIndex        =   42
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "Installation Notes"
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
      Left            =   240
      TabIndex        =   47
      Top             =   5400
      Width           =   2412
   End
End
Attribute VB_Name = "frmEditComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Data are only saved back to the active truck upon leaving the screen.  This allows the "Cancel"
'option to work w/o having to undo changes.

Private sLenDispFormat As String
Private CompNdx As Integer
Private cLocalComponent As clsComponent 'Locally-edited Component object
Private bSaveChanges As Boolean

Public Sub EditComponent(ComponentIndex As Integer, frmParent As Form)
    'Called from elsewhere to edit the indexed component in the global cActiveTruck object
    
    Set cLocalComponent = New clsComponent
    Set cLocalComponent = ComponentCopy(cActiveTruck.Components(ComponentIndex))
    bSaveChanges = False
    CompNdx = ComponentIndex
    InitForm
    Me.Show vbModal, frmParent
    
    If bSaveChanges Then
        cActiveTruck.Components.Remove (ComponentIndex)
        cActiveTruck.Components.Add cLocalComponent
    End If
    Set cLocalComponent = Nothing
End Sub


Private Sub InitForm()
    Dim i%
    Dim dblTemp As Double
    Dim vLabel
    On Error GoTo errHandler
    
    'Set formatting filter for distance
    sLenDispFormat = "###0.000 "
    If Int(cGlobalInfo.DistanceUnits.Multiplier * 100) = 3937 Then
        'Inches
        sLenDispFormat = "###0.0 "
    End If
    
    'Add Proper Engineering Units
    For Each vLabel In lblDistanceUnits
        vLabel.Caption = cGlobalInfo.DistanceUnits.Display
    Next
    For Each vLabel In lblMassUnits
        vLabel.Caption = cGlobalInfo.MassUnits.Display
    Next
    For Each vLabel In lblVolumeUnits
        vLabel.Caption = cGlobalInfo.VolumeUnits.Display
    Next

    'Full Name
    txtFullName.Text = cLocalComponent.FullName
    'Display Name
    txtDisplayName.Text = cLocalComponent.DisplayName
    'Offset
    If cLocalComponent.LocationReference = orFrontAxle Then
        'Make dist. to center of single imaginary front axle into
        ' distance from actual front axle
        dblTemp = cLocalComponent.Offset + cActiveTruck.Chassis.TwinSteerSeparation / 2
        dblTemp = dblTemp * cGlobalInfo.DistanceUnits.Multiplier
    Else
        dblTemp = cLocalComponent.Offset * cGlobalInfo.DistanceUnits.Multiplier
    End If
    txtCompOffset.Text = Format(dblTemp, sLenDispFormat)
    'Empty Weight
    dblTemp = cLocalComponent.EmptyWeight * cGlobalInfo.MassUnits.Multiplier
    txtEmptyWeight.Text = Format(dblTemp, "#0")
    'Empty CG Location
    dblTemp = cLocalComponent.EmptyCG * cGlobalInfo.DistanceUnits.Multiplier
    txtEmptyCG.Text = Format(dblTemp, sLenDispFormat)
    'Installation Notes
    txtNotes.Text = cLocalComponent.InstallationNotes
    
    'Contents Type
    cboContentsType.Clear
    cboContentsType.AddItem "None", ctNone
    cboContentsType.AddItem "Fuel", ctFuel
    cboContentsType.AddItem "Water", ctWater
    cboContentsType.AddItem "GasA", ctGasA
    cboContentsType.AddItem "GasB", ctGasB
    cboContentsType.AddItem "Additive", ctAdditive
    cboContentsType.AddItem "Other", ctOther
    cboContentsType.ListIndex = cLocalComponent.ContentsType

    'Location Reference
    cboLocationReference.Clear
    cboLocationReference.AddItem "Body", orBodyOrigin
    cboLocationReference.AddItem "Chassis", orFrontAxle
    cboLocationReference.ListIndex = cLocalComponent.LocationReference
    
    'PlacementAllowable
    cboPlacementAllowable.Clear
    cboPlacementAllowable.AddItem "Either Side", paEitherSide
    cboPlacementAllowable.AddItem "StreetSideStd", paStreetSideStd
    cboPlacementAllowable.AddItem "CurbSideStd", paCurbSideStd
    cboPlacementAllowable.AddItem "Center", paCenter
    cboPlacementAllowable.ListIndex = cLocalComponent.PlacementAllowable
    
    'Placement
    cboPlacement.Clear
    cboPlacement.AddItem "Not Placed", plNotPlaced
    cboPlacement.AddItem "StreetSideStd", plStreetSideStd
    cboPlacement.AddItem "CurbSideStd", plCurbSideStd
    cboPlacement.AddItem "Center", plCenter
    cboPlacement.ListIndex = cLocalComponent.Placement
    
    'Drawing File
    txtStreetSideStd.Text = cLocalComponent.StreetSideStd
    txtCurbSideStd.Text = cLocalComponent.CurbSideStd
    
    'Volume Capacity (if applicable)
    'If cLocalComponent.Capacity.Volume > 0 Then
        dblTemp = cLocalComponent.Capacity.Volume * cGlobalInfo.VolumeUnits.Multiplier
        txtVolume.Text = Format(dblTemp, "#0")
    'End If
    
    'Default Weight (if applicable)
    If cLocalComponent.Capacity.DefaultWtContents <> "" Then
        dblTemp = CDbl(cLocalComponent.Capacity.DefaultWtContents) _
                 * cGlobalInfo.MassUnits.Multiplier
        txtDefaultWtContents.Text = Format(dblTemp, "#0")
    Else
        txtDefaultWtContents.Text = "N/A"
        txtDefaultWtContents.Enabled = False
    End If
    
    'Default Volume (if applicable)
    If cLocalComponent.Capacity.DefaultVolContents <> "" Then
        dblTemp = CDbl(cLocalComponent.Capacity.DefaultVolContents) _
                 * cGlobalInfo.VolumeUnits.Multiplier
        txtDefaultVolContents.Text = Format(dblTemp, "#0")
    Else
        txtDefaultVolContents.Text = "N/A"
        txtDefaultVolContents.Enabled = False
    End If
    
    'Density Contents
    If cLocalComponent.ContentsType = ctOther Then
        dblTemp = CDbl(cLocalComponent.Capacity.DensityContents)
        txtDensityContents.Text = Format(dblTemp, "#0.00")
    Else
        txtDensityContents.Text = "N/A"
        txtDensityContents.Enabled = False
    End If
    
    'Uses Sight Gauge
    If cLocalComponent.Capacity.UsesSightGauge Then
        chkUsesSightGauge.Value = vbChecked
    Else
        chkUsesSightGauge.Value = vbUnchecked
    End If
    
    'Relationships
    Dim cRelationship As clsFillRelationship
    chkAN.Value = vbUnchecked
    chkEmulsion.Value = vbUnchecked
    For i% = 0 To 1
        txtMultiplier(i%).Text = ""
        txtOffset(i%).Text = ""
    Next i%
    dblTemp = 0#
    For Each cRelationship In cLocalComponent.FillRelationShips
        If cRelationship.ParentProduct = ptAN Then
            i% = 0
            chkAN.Value = vbChecked
        Else:
            i% = 1
            chkEmulsion.Value = vbChecked
        End If
        txtMultiplier(i%).Text = cRelationship.Multiplier
        txtOffset(i%).Text = cRelationship.Offset
    Next
    
    'Protect certain users from themselves
    cmdSave.Enabled = bCanCreateALObjects
    
    Exit Sub
errHandler:
    ErrorIn "frmEditComponent.InitForm"
End Sub


Private Sub cboContentsType_Click()
    'Default Wt and Vol properties ONLY applicable where Contents = "Other"
    'Density contents n/a for "None" and filled from Global Properties if fuel, gas, water, etc.
    If cboContentsType.ListIndex = ctOther Then
        txtDefaultWtContents.Enabled = True
        txtDefaultVolContents.Enabled = True
        txtDensityContents.Enabled = True
    
        fraAN.Enabled = True
        fraEmulsion.Enabled = True
        cmdStickLength.Enabled = True
        cmdContentCG.Enabled = True
        chkUsesSightGauge.Enabled = True
        txtVolume.Enabled = True
    ElseIf cboContentsType.ListIndex = ctNone Then
        fraAN.Enabled = False
        fraEmulsion.Enabled = False
        cmdStickLength.Enabled = False
        cmdContentCG.Enabled = False
        chkUsesSightGauge.Enabled = False
        txtVolume.Enabled = False
        txtDensityContents.Enabled = False
        txtDefaultVolContents.Enabled = False
        txtDefaultWtContents.Enabled = False
    Else
        txtDefaultWtContents.Enabled = False
        txtDefaultVolContents.Enabled = False
        txtDensityContents.Enabled = False
    
        fraAN.Enabled = True
        fraEmulsion.Enabled = True
        cmdStickLength.Enabled = True
        cmdContentCG.Enabled = True
        chkUsesSightGauge.Enabled = True
        txtVolume.Enabled = True
    End If
End Sub

Private Sub cboContentsType_Validate(Cancel As Boolean)
    cboContentsType_Click
End Sub


Private Sub cboPlacement_Validate(Cancel As Boolean)
    Dim plNow As PlacementLocation
    Dim sMsg As String
    Dim bOK As Boolean
    
    plNow = cboPlacement.ListIndex
    bOK = True
    Select Case cboPlacementAllowable.ListIndex
    Case paEitherSide
        If plNow = plCurbSideStd Or plNow = plStreetSideStd Then
            'Leave alone
        Else
            bOK = False
        End If
    Case paStreetSideStd
        If plNow <> plStreetSideStd Then
            bOK = False 'Wrong side
        End If
    Case paCurbSideStd
        If plNow <> plCurbSideStd Then
            bOK = False 'Wrong side
        End If
    Case paCenter
        If plNow <> plCenter Then
            bOK = False 'Wrong side
        End If
    End Select
    
    If Not bOK Then
        sMsg = "This placement option is not allowed." & vbCrLf & _
               "Select another Placement option or change the 'Placement Allowable' option."
        MsgBox sMsg, vbExclamation
        Cancel = True
    End If
End Sub

Private Sub cboPlacementAllowable_Validate(Cancel As Boolean)
    'Make sure .Placement doesn't conflict
    cboPlacement.ListIndex = SetDefaultPlacement(cboPlacementAllowable.ListIndex, cboPlacement.ListIndex)
End Sub

Private Sub chkAN_Click()
    txtMultiplier(0).Enabled = (chkAN.Value = vbChecked)
    txtOffset(0).Enabled = (chkAN.Value = vbChecked)
End Sub

Private Sub chkEmulsion_Click()
    txtMultiplier(1).Enabled = (chkEmulsion.Value = vbChecked)
    txtOffset(1).Enabled = (chkEmulsion.Value = vbChecked)
End Sub

Private Sub cmdContentCG_Click()
    Dim colCoeff As New Collection
    Dim sMsg As String
    On Error GoTo errHandler
    
    Set colCoeff = cLocalComponent.Capacity.ContentCG
    
    If colCoeff.Count = 0 Then
        sMsg = "This component does not currently have any 'ContentCG' data." & vbCrLf & _
               "Do you really want to add this information?"
        If MsgBox(sMsg, vbYesNo, "Add 'ContentCG' Coefficients?") = vbNo Then
            Exit Sub
        End If
    End If
    
    Set cLocalComponent.Capacity.ContentCG = frmEditCoefficients.EditCoefficients(colCoeff, Me)
    Exit Sub
errHandler:
    ErrorIn "frmEditComponent.cmdContentCG_Click"
End Sub


Private Sub cmdStickLength_Click()
    Dim colCoeff As New Collection
    Dim sMsg As String
    On Error GoTo errHandler
    
    Set colCoeff = cLocalComponent.Capacity.StickLength
    
    If colCoeff.Count = 0 Then
        sMsg = "This component does not currently have any 'StickLength' data." & vbCrLf & _
               "Do you really want to add this information?"
        If MsgBox(sMsg, vbYesNo, "Add 'StickLength' Coefficients?") = vbNo Then
            Exit Sub
        End If
    End If
    
    Set cLocalComponent.Capacity.StickLength = frmEditCoefficients.EditCoefficients(colCoeff, Me)
    Exit Sub
errHandler:
    ErrorIn "frmEditComponent.cmdStickLength_Click"
End Sub

Private Function ComponentProblems(cComp As clsComponent) As String
    'This function checks the edited component to see if it passes all the
    ' rules for a valid component.  Errors are returned.
    Dim sTemp As String
    Dim dblTemp  As Double
    Dim cRelationship As clsFillRelationship
    On Error GoTo errHandler
    
    ComponentProblems = "Unforseen error in frmEditComponent.ComponentProblems()" 'Default = fail
    
    'Full Name
    cLocalComponent.FullName = txtFullName.Text
    'Display Name
    cLocalComponent.DisplayName = txtDisplayName.Text
    'Offset
    If cLocalComponent.LocationReference = orFrontAxle Then
        'Correct measurement so distance from axle is really
        ' dist. to center of single imaginary front axle
        dblTemp = CDbl(txtCompOffset.Text) / cGlobalInfo.DistanceUnits.Multiplier
        cLocalComponent.Offset = dblTemp - cActiveTruck.Chassis.TwinSteerSeparation / 2
    Else
        cLocalComponent.Offset = CDbl(txtCompOffset.Text) / cGlobalInfo.DistanceUnits.Multiplier
    End If
    'Empty Weight
    cLocalComponent.EmptyWeight = CDbl(txtEmptyWeight.Text) / cGlobalInfo.MassUnits.Multiplier
    'Empty CG Location
    cLocalComponent.EmptyCG = CDbl(txtEmptyCG.Text) / cGlobalInfo.DistanceUnits.Multiplier
    'Installation Notes
    cLocalComponent.InstallationNotes = txtNotes.Text
    'Contents Type
    cLocalComponent.ContentsType = cboContentsType.ListIndex
    'Location Reference
    cLocalComponent.LocationReference = cboLocationReference.ListIndex
    
    'PlacementAllowable
    cLocalComponent.PlacementAllowable = cboPlacementAllowable.ListIndex
    'Placement
    cLocalComponent.Placement = cboPlacement.ListIndex
    'StreetSideStd
    cLocalComponent.StreetSideStd = txtStreetSideStd.Text
    'CurbSideStd
    cLocalComponent.CurbSideStd = txtCurbSideStd.Text
    
    'Volume Capacity
    sTemp = Trim$(txtVolume.Text)
    cLocalComponent.Capacity.Volume = CDbl(sTemp) / cGlobalInfo.VolumeUnits.Multiplier
    
    'Default Weight (if applicable)
    cLocalComponent.Capacity.DefaultWtContents = ""
    If cLocalComponent.ContentsType = ctOther Then
        sTemp = UCase$(Trim$(txtDefaultWtContents.Text))
        If IsNumber(sTemp) Then
            If CDbl(sTemp) <> 0 Then
                'Only save a value if Contents=Other and value is a non-zero number
                cLocalComponent.Capacity.DefaultWtContents = CStr(CDbl(sTemp) / cGlobalInfo.MassUnits.Multiplier)
            End If
        End If
    End If
    
    'Default Volume (if applicable)
    cLocalComponent.Capacity.DefaultVolContents = ""
    If cLocalComponent.ContentsType = ctOther Then
        sTemp = UCase$(Trim$(txtDefaultVolContents.Text))
        If IsNumber(sTemp) Then
            If CDbl(sTemp) <> 0 Then
                'Only save a value if Contents=Other and value is a non-zero number
                cLocalComponent.Capacity.DefaultVolContents = CStr(CDbl(sTemp) / cGlobalInfo.VolumeUnits.Multiplier)
            End If
        End If
    End If
    
    'Density Contents
    Select Case cLocalComponent.ContentsType
    Case ctNone
        cLocalComponent.Capacity.DensityContents = 0
    Case ctFuel
        cLocalComponent.Capacity.DensityContents = cGlobalInfo.DensityFuel
    Case ctWater
        cLocalComponent.Capacity.DensityContents = cGlobalInfo.DensityWater
    Case ctGasA
        cLocalComponent.Capacity.DensityContents = cGlobalInfo.DensityGasA
    Case ctGasB
        cLocalComponent.Capacity.DensityContents = cGlobalInfo.DensityGasB
    Case ctAdditive
        cLocalComponent.Capacity.DensityContents = cGlobalInfo.DensityAdditive
    Case ctOther
        If txtDensityContents.Enabled And IsNumber(txtDensityContents.Text) Then
            cLocalComponent.Capacity.DensityContents = CDbl(txtDensityContents.Text)
        End If
    End Select
    
    
    'Uses Sight Gauge
    cLocalComponent.Capacity.UsesSightGauge = False
    If cLocalComponent.ContentsType <> ctNone Then
        cLocalComponent.Capacity.UsesSightGauge = (chkUsesSightGauge.Value = vbChecked)
    End If
    
    'Relationships
    Set cLocalComponent.FillRelationShips = New clsFillRelationships
    If cLocalComponent.ContentsType <> ctNone Then
        If chkAN.Value = vbChecked Then
            If IsNumber(txtMultiplier(0).Text) And IsNumber(txtOffset(0).Text) Then
                Set cRelationship = New clsFillRelationship
                cRelationship.Multiplier = CDbl(txtMultiplier(0).Text)
                cRelationship.Offset = CDbl(txtOffset(0).Text)
                cRelationship.ParentProduct = ptAN
                cLocalComponent.FillRelationShips.Add cRelationship
            Else
                ComponentProblems = "You must enter valid numbers for AN Relationship Offset and Multiplier."
                GoTo CleanUp
            End If
        End If
        If chkEmulsion.Value = vbChecked Then
            If IsNumber(txtMultiplier(1).Text) And IsNumber(txtOffset(1).Text) Then
                Set cRelationship = New clsFillRelationship
                cRelationship.Multiplier = CDbl(txtMultiplier(1).Text)
                cRelationship.Offset = CDbl(txtOffset(1).Text)
                cRelationship.ParentProduct = ptEmulsion
                cLocalComponent.FillRelationShips.Add cRelationship
            Else
                ComponentProblems = "You must enter valid numbers for Emulsion Relationship Offset and Multiplier."
                GoTo CleanUp
            End If
        End If
    End If
    
    'Stick Length and ContentCG were defined elsewhere.  Delete if N/A.
    If cLocalComponent.ContentsType = ctNone Then
        Set cLocalComponent.Capacity.StickLength = New Collection
        Set cLocalComponent.Capacity.ContentCG = New Collection
    End If
    
    ComponentProblems = "" 'No errors if we got this far
CleanUp:
    Set cRelationship = Nothing
    Exit Function
errHandler:
    Resume CleanUp
End Function

Private Sub cmdExit_Click()
    'Don't save changes, just exit
    bSaveChanges = False
    Unload Me
End Sub


Private Sub cmdOK_Click()
    'Save changes to truck
    Dim sMsg As String
    
    'Read data on form into cLocalComponent.  Give error msg if appl.
    sMsg = ComponentProblems(cLocalComponent)
    If sMsg <> "" Then
        MsgBox sMsg, vbExclamation, "Cannot Save Component"
        Exit Sub
    End If
    
    bSaveChanges = True
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sMsg As String
    Dim sFile As String
    
    'Read data on form into cLocalComponent.  Give error msg if appl.
    sMsg = ComponentProblems(cLocalComponent)
    If sMsg <> "" Then
        MsgBox sMsg, vbExclamation, "Cannot Save Component"
        Exit Sub
    End If
    
    sMsg = SaveComponentObject(cLocalComponent)
    If sMsg <> "" Then
        MsgBox sMsg, vbExclamation, "Problem Saving Component"
    End If

End Sub

Private Sub txtCompOffset_Validate(Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtCompOffset.Text) Then
        sMsg = "You must enter a valid number for the component Offset"
        MsgBox sMsg, vbExclamation, "Invalid Entry"
        Cancel = True
    End If
End Sub


Private Sub txtEmptyCG_Validate(Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtEmptyCG.Text) Then
        sMsg = "You must enter a valid number for (Empty) CG Location"
        MsgBox sMsg, vbExclamation, "Invalid Entry"
        Cancel = True
    End If
End Sub

Private Sub txtEmptyWeight_Validate(Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtEmptyWeight.Text) Then
        sMsg = "You must enter a valid number for Empty Weight"
        MsgBox sMsg, vbExclamation, "Invalid Entry"
        Cancel = True
    End If
End Sub


Private Sub txtVolume_Validate(Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtVolume.Text) Then
        sMsg = "You must enter a valid number for Capacity (Volume)"
        MsgBox sMsg, vbExclamation, "Invalid Entry"
        Cancel = True
    End If
End Sub


Private Sub txtDensityContents_Validate(Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtDensityContents.Text) Then
        sMsg = "You must enter a valid number for Specific Gravity"
        MsgBox sMsg, vbExclamation, "Invalid Entry"
        Cancel = True
    End If
End Sub

Private Sub txtMultiplier_Validate(Index As Integer, Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtMultiplier(Index).Text) Then
        sMsg = "You must enter a valid number for Multiplier"
        MsgBox sMsg, vbExclamation, "Invalid Entry"
        Cancel = True
    End If
End Sub


Private Sub txtOffset_Validate(Index As Integer, Cancel As Boolean)
    Dim sMsg As String
    
    If Not IsNumber(txtOffset(Index).Text) Then
        sMsg = "You must enter a valid number for Offset"
        MsgBox sMsg, vbExclamation, "Invalid Entry"
        Cancel = True
    End If
End Sub


