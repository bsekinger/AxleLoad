VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmEditChassis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Chassis"
   ClientHeight    =   8472
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   7812
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8472
   ScaleWidth      =   7812
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraObject 
      Height          =   6252
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7572
      Begin C1SizerLibCtl.C1Tab tabTagDisplay 
         Height          =   2652
         Left            =   600
         TabIndex        =   13
         Top             =   3240
         Width           =   2652
         _cx             =   4678
         _cy             =   4678
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   800
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   14737632
         TabOutlineColor =   0
         FrontTabForeColor=   -2147483630
         Caption         =   "Tag 1"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   2
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   0   'False
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   4
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin VB.PictureBox picPage 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2364
            Left            =   12
            ScaleHeight     =   2364
            ScaleWidth      =   2628
            TabIndex        =   14
            Top             =   276
            Width           =   2628
            Begin VB.TextBox txtTagWtLimit 
               Height          =   315
               Left            =   120
               TabIndex        =   19
               ToolTipText     =   "Manufacturer's specified loading limit"
               Top             =   1836
               Width           =   852
            End
            Begin VB.TextBox txtTagWeight 
               Height          =   315
               Left            =   120
               TabIndex        =   18
               ToolTipText     =   "Normally in the range of 1600 lbs."
               Top             =   1116
               Width           =   852
            End
            Begin VB.TextBox txtTagLoc 
               Height          =   315
               Left            =   120
               TabIndex        =   17
               ToolTipText     =   "Distance from frontmost axle to tag."
               Top             =   396
               Width           =   852
            End
            Begin VB.PictureBox picNew 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               DrawMode        =   1  'Blackness
               ForeColor       =   &H80000008&
               Height          =   384
               Left            =   2160
               Picture         =   "frmEditChassis.frx":0000
               ScaleHeight     =   384
               ScaleWidth      =   384
               TabIndex        =   16
               ToolTipText     =   "Add another Tag"
               Top             =   120
               Width           =   384
            End
            Begin VB.PictureBox picDelete 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   384
               Left            =   2160
               Picture         =   "frmEditChassis.frx":030A
               ScaleHeight     =   384
               ScaleWidth      =   384
               TabIndex        =   15
               ToolTipText     =   "Remove this tag"
               Top             =   1920
               Width           =   384
            End
            Begin VB.Label lblMassUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "Mass"
               Height          =   276
               Index           =   1
               Left            =   1080
               TabIndex        =   25
               Top             =   1956
               Width           =   732
            End
            Begin VB.Label lblMassUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "Mass"
               Height          =   276
               Index           =   0
               Left            =   1080
               TabIndex        =   24
               Top             =   1236
               Width           =   732
            End
            Begin VB.Label lblDistanceUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "DistUnits"
               Height          =   276
               Index           =   0
               Left            =   1080
               TabIndex        =   23
               Top             =   516
               Width           =   732
            End
            Begin VB.Label lblTagLoc 
               BackStyle       =   0  'Transparent
               Caption         =   "Tag 1"
               Height          =   276
               Index           =   0
               Left            =   120
               TabIndex        =   22
               Top             =   120
               Width           =   1932
            End
            Begin VB.Label lblTagWeight 
               BackStyle       =   0  'Transparent
               Caption         =   "Weight of Tag Assy"
               Height          =   276
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   840
               Width           =   1572
            End
            Begin VB.Label lblTagWtLimit 
               BackStyle       =   0  'Transparent
               Caption         =   "Loading Limit"
               Height          =   276
               Index           =   0
               Left            =   120
               TabIndex        =   20
               Top             =   1596
               Width           =   1212
            End
         End
      End
      Begin VB.ListBox lstComponents 
         Height          =   1776
         Left            =   5640
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   4080
         Width           =   1692
      End
      Begin VSFlex8LCtl.VSFlexGrid grdChassis 
         Height          =   3012
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   7212
         _cx             =   12721
         _cy             =   5313
         Appearance      =   1
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   17
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmEditChassis.frx":074C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   6
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblTags 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditChassis.frx":0899
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2292
         Left            =   180
         TabIndex        =   12
         Top             =   3600
         Width           =   348
      End
      Begin VB.Label lblComponents 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "If saving as a new chassis file, Attach these Components:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   5640
         TabIndex        =   6
         Top             =   3360
         Width           =   1692
      End
   End
   Begin VB.Frame fraObjectFile 
      Caption         =   "Descriptions for New Chassis Object File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   7572
      Begin VB.TextBox txtObjectFullName 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         ToolTipText     =   "Only used when saving a New Chassis File."
         Top             =   720
         Width           =   6132
      End
      Begin VB.TextBox txtObjectDisplayName 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "Only used when saving a New Chassis File."
         Top             =   360
         Width           =   2532
      End
      Begin VB.Label lblObjectFullName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Description:"
         Height          =   276
         Left            =   0
         TabIndex        =   10
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label lblObjectDisplayName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Display Name:"
         Height          =   276
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   1212
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   612
      Left            =   120
      TabIndex        =   2
      Top             =   7764
      Width           =   1572
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   $"frmEditChassis.frx":08B5
      Height          =   612
      Left            =   3240
      TabIndex        =   1
      Top             =   7764
      Width           =   1572
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   $"frmEditChassis.frx":08D3
      Height          =   612
      Left            =   6120
      TabIndex        =   0
      Top             =   7764
      Width           =   1572
   End
End
Attribute VB_Name = "frmEditChassis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Data are only saved back to the active truck upon leaving the screen.  This allows the "Cancel"
'option to work w/o having to undo changes.

Private sLenDispFormat As String
Private cTags As clsTags
Private LastTagIndex As Long

Private Sub Form_Load()
    txtObjectDisplayName.Text = ""
    InitForm
End Sub

Private Sub InitForm()
    Dim i%
    Dim dblTemp As Double
    Dim vLabel
    Dim sTemp As String
        
    'Set formatting filter for distance
    sLenDispFormat = "###0.00 "
    If Int(cGlobalInfo.DistanceUnits.Multiplier * 100) = 3937 Then
        'Inches
        sLenDispFormat = "###0.0 "
    End If
    'Caption of Form
    If Len(cActiveTruck.SN) > 2 Then
        If Mid$(cActiveTruck.SN, 1, 1) = "q" Then
            sTemp = "(Q" & Mid$(cActiveTruck.SN, 2) & ")"
        ElseIf UCase$(Mid$(cActiveTruck.SN, 1, 2)) = "SN" Then
            sTemp = "(SN" & Mid$(cActiveTruck.SN, 3) & ")"
        Else
            sTemp = "(" & Trim$(cActiveTruck.SN) & ")"
        End If
    Else
        sTemp = ""
    End If
    Me.Caption = "Edit Chassis" & sTemp
    'Add Proper Engineering Units
    For Each vLabel In lblDistanceUnits
        vLabel.Caption = cGlobalInfo.DistanceUnits.Display
    Next
    For Each vLabel In lblMassUnits
        vLabel.Caption = cGlobalInfo.MassUnits.Display
    Next
    
    grdChassis.ColAlignment(1) = flexAlignRightCenter
    grdChassis.ColAlignment(2) = flexAlignLeftCenter
    
    'FullName
    grdChassis.Cell(flexcpData, 0, 1) = "FullName"
    grdChassis.Cell(flexcpText, 0, 1) = cActiveTruck.Chassis.FullName
    grdChassis.Cell(flexcpAlignment, 0, 1) = flexAlignGeneral
    grdChassis.Cell(flexcpText, 0, 2) = ""
    grdChassis.Cell(flexcpAlignment, 0, 1) = flexAlignGeneral

    'DisplayName
    grdChassis.Cell(flexcpData, 1, 1) = "DisplayName"
    grdChassis.Cell(flexcpText, 1, 1) = cActiveTruck.Chassis.DisplayName
    grdChassis.Cell(flexcpAlignment, 1, 1) = flexAlignGeneral
    grdChassis.Cell(flexcpText, 1, 2) = ""
    grdChassis.Cell(flexcpAlignment, 1, 2) = flexAlignGeneral
    
    'WB
    grdChassis.Cell(flexcpData, 2, 1) = "WB"
    dblTemp = (cActiveTruck.Chassis.WB + cActiveTruck.Chassis.TwinSteerSeparation / 2) * cGlobalInfo.DistanceUnits.Multiplier
    grdChassis.Cell(flexcpText, 2, 1) = Format(dblTemp, sLenDispFormat)
    grdChassis.Cell(flexcpText, 2, 2) = cGlobalInfo.DistanceUnits.Display
    
    'BackOfCab
    grdChassis.Cell(flexcpData, 3, 1) = "BackOfCab"
    dblTemp = cActiveTruck.Chassis.BackOfCab * cGlobalInfo.DistanceUnits.Multiplier
    grdChassis.Cell(flexcpText, 3, 1) = Format(dblTemp, sLenDispFormat)
    grdChassis.Cell(flexcpText, 3, 2) = cGlobalInfo.DistanceUnits.Display
    
    'WtFront
    grdChassis.Cell(flexcpData, 4, 1) = "WtFront"
    dblTemp = cActiveTruck.Chassis.WtFront * cGlobalInfo.MassUnits.Multiplier
    grdChassis.Cell(flexcpText, 4, 1) = Format(dblTemp, "#")
    grdChassis.Cell(flexcpText, 4, 2) = cGlobalInfo.MassUnits.Display
    
    'WtRear
    grdChassis.Cell(flexcpData, 5, 1) = "WtRear"
    dblTemp = cActiveTruck.Chassis.WtRear * cGlobalInfo.MassUnits.Multiplier
    grdChassis.Cell(flexcpText, 5, 1) = Format(dblTemp, "#")
    grdChassis.Cell(flexcpText, 5, 2) = cGlobalInfo.MassUnits.Display
    
    'WtLimitFront
    grdChassis.Cell(flexcpData, 6, 1) = "WtLimitFront"
    dblTemp = cActiveTruck.Chassis.WtLimitFront * cGlobalInfo.MassUnits.Multiplier
    grdChassis.Cell(flexcpText, 6, 1) = Format(dblTemp, "#")
    grdChassis.Cell(flexcpText, 6, 2) = cGlobalInfo.MassUnits.Display
    
    'WtLimitRear
    grdChassis.Cell(flexcpData, 7, 1) = "WtLimitRear"
    dblTemp = cActiveTruck.Chassis.WtLimitRear * cGlobalInfo.MassUnits.Multiplier
    grdChassis.Cell(flexcpText, 7, 1) = Format(dblTemp, "#")
    grdChassis.Cell(flexcpText, 7, 2) = cGlobalInfo.MassUnits.Display
    
    'WtLimitTotal
    grdChassis.Cell(flexcpData, 8, 1) = "WtLimitTotal"
    dblTemp = cActiveTruck.Chassis.WtLimitTotal * cGlobalInfo.MassUnits.Multiplier
    grdChassis.Cell(flexcpText, 8, 1) = Format(dblTemp, "#")
    grdChassis.Cell(flexcpText, 8, 2) = cGlobalInfo.MassUnits.Display
    
    'TandemSpacing
    grdChassis.Cell(flexcpData, 9, 1) = "TandemSpacing"
    dblTemp = cActiveTruck.Chassis.TandemSpacing * cGlobalInfo.DistanceUnits.Multiplier
    grdChassis.Cell(flexcpText, 9, 1) = Format(dblTemp, sLenDispFormat)
    grdChassis.Cell(flexcpText, 9, 2) = cGlobalInfo.DistanceUnits.Display
    
    'PlacementAllowable
    grdChassis.Cell(flexcpAlignment, 10, 1) = flexAlignGeneral
    grdChassis.Cell(flexcpData, 10, 1) = "PlacementAllowable"
    grdChassis.Cell(flexcpText, 10, 1) = "Center" 'cActiveTruck.Chassis.PlacementAllowableString
    
    'StreetSideStd
    grdChassis.Cell(flexcpAlignment, 11, 1) = flexAlignGeneral
    grdChassis.Cell(flexcpData, 11, 1) = "StreetSideStd"
    grdChassis.Cell(flexcpText, 11, 1) = cActiveTruck.Chassis.StreetSideStd

    'CurbSideStd
    grdChassis.Cell(flexcpAlignment, 12, 1) = flexAlignGeneral
    grdChassis.Cell(flexcpData, 12, 1) = "CurbSideStd"
    grdChassis.Cell(flexcpText, 12, 1) = cActiveTruck.Chassis.CurbSideStd
    
    'WheelDia
    grdChassis.Cell(flexcpData, 13, 1) = "WheelDia"
    dblTemp = cActiveTruck.Chassis.WheelDia * cGlobalInfo.DistanceUnits.Multiplier
    grdChassis.Cell(flexcpText, 13, 1) = Format(dblTemp, sLenDispFormat)
    grdChassis.Cell(flexcpText, 13, 2) = cGlobalInfo.DistanceUnits.Display
    
    'WheelY
    grdChassis.Cell(flexcpData, 14, 1) = "WheelY"
    dblTemp = cActiveTruck.Chassis.WheelY * cGlobalInfo.DistanceUnits.Multiplier
    grdChassis.Cell(flexcpText, 14, 1) = Format(dblTemp, sLenDispFormat)
    grdChassis.Cell(flexcpText, 14, 2) = cGlobalInfo.DistanceUnits.Display
    
    'FrameRailHt
    grdChassis.Cell(flexcpData, 15, 1) = "FrameRailHt"
    dblTemp = cActiveTruck.Chassis.FrameRailHt * cGlobalInfo.DistanceUnits.Multiplier
    grdChassis.Cell(flexcpText, 15, 1) = Format(dblTemp, sLenDispFormat)
    grdChassis.Cell(flexcpText, 15, 2) = cGlobalInfo.DistanceUnits.Display
    
    'TwinSteerSeparation
    grdChassis.Cell(flexcpData, 16, 1) = "TwinSteerSeparation"
    dblTemp = cActiveTruck.Chassis.TwinSteerSeparation * cGlobalInfo.DistanceUnits.Multiplier
    grdChassis.Cell(flexcpText, 16, 1) = Format(dblTemp, sLenDispFormat)
    grdChassis.Cell(flexcpText, 16, 2) = cGlobalInfo.DistanceUnits.Display
    
    'Set Tag data
    InitTags tabTagDisplay
    
    'Show components
    lstComponents.Clear
    For i% = 1 To cActiveTruck.Components.Count
        lstComponents.AddItem cActiveTruck.Components(i%).DisplayName, i% - 1
    Next i%
    
    'Protect certain users from themselves
    lstComponents.Enabled = bCanCreateALObjects
    txtObjectDisplayName.Enabled = bCanCreateALObjects
    txtObjectFullName.Enabled = bCanCreateALObjects
    cmdSave.Enabled = bCanCreateALObjects
End Sub

Private Sub grdChassis_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static LastRow As Long
    Dim CurRow As Long

    ' get coordinates
    CurRow = grdChassis.MouseRow

    ' update tooltip text
    If CurRow <> LastRow Then
        'only update tooltiptext if row has changed
        Select Case CurRow
        Case 2
            'WB
            grdChassis.ToolTipText = "Distance from front-most axle to center of rear tandem."
        Case 8
            'WtLimitTotal
            grdChassis.ToolTipText = "If WtLimitTotal=0, program will set Total = Front + Rear limits."
        Case 9
            'Tandem spacing
            grdChassis.ToolTipText = "Normally 52in for tandems. TandemSpacing=0 means single rear axle."
        Case 10
            'PlacementAllowable
            grdChassis.ToolTipText = "The chassis is always in the 'Center' position."
        Case 11
            'StreetSideStd
            grdChassis.ToolTipText = "DXF file that shows Street-Side view of the chassis"
        Case 12
            'CurbSideStd
            grdChassis.ToolTipText = "DXF file that shows Curb-Side view of the chassis"
        Case 13
            'WheelDia
            grdChassis.ToolTipText = "Wheel Diameter.  Used only in making layout dwg."
        Case 14
            'WheelY
            grdChassis.ToolTipText = "Dist. fm frame rail top to wheel center.  Used only in making layout dwg."
        Case 15
            'FrameRailHt
            grdChassis.ToolTipText = "Height of frame rail.  Used only in making layout dwg."
        Case 16
            'TwinSteerSeparation
            grdChassis.ToolTipText = "(Default=0) If Dual-Steer chassis, this is distance b/w front axles."
        Case Else
           grdChassis.ToolTipText = ""
        End Select
        LastRow = CurRow
    End If
End Sub

Private Sub grdChassis_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'User made a change, so we should validate the change.
    Dim strErr As String
    
    If Col = 2 Or Col = 0 Then
        'user shouldn't be able to change this column
        Cancel = True
        Exit Sub
    End If
    
    strErr = strErr & VerifyCell(grdChassis.EditText, grdChassis.Cell(flexcpData, Row, Col))
    
    If strErr <> "" Then
        MsgBox strErr, vbExclamation, "Warning"
        Cancel = True
        Exit Sub
    End If
End Sub


Private Sub txtTagLoc_Validate(Cancel As Boolean)
    Dim strErr As String
    
    strErr = TagLocErr(tabTagDisplay.CurrTab + 1)
    If strErr <> "" Then
        MsgBox strErr, vbExclamation, "Warning"
        Cancel = True
        Exit Sub
    End If
    cTags(tabTagDisplay.CurrTab + 1).Location = CDbl(txtTagLoc.Text) / cGlobalInfo.DistanceUnits.Multiplier
End Sub

Private Sub txtTagWeight_Validate(Cancel As Boolean)
    Dim strErr As String
    
    strErr = TagWeightErr(tabTagDisplay.CurrTab + 1)
    If strErr <> "" Then
        MsgBox strErr, vbExclamation, "Warning"
        Cancel = True
        Exit Sub
    End If
    cTags(tabTagDisplay.CurrTab + 1).Weight = CDbl(txtTagWeight.Text) / cGlobalInfo.MassUnits.Multiplier
End Sub

Private Sub txtTagWtLimit_Validate(Cancel As Boolean)
    Dim strErr As String
    
    strErr = TagWtLimitErr(tabTagDisplay.CurrTab + 1)
    If strErr <> "" Then
        MsgBox strErr, vbExclamation, "Warning"
        Cancel = True
        Exit Sub
    End If
    cTags(tabTagDisplay.CurrTab + 1).WtLimit = CDbl(txtTagWtLimit.Text) / cGlobalInfo.MassUnits.Multiplier
End Sub

Private Function TagLocErr(Index As Integer) As String
    'Returns an error msg if the indicated 'txtTagLoc()' control contains bad data
    Dim strText As String
    Dim a$
    
    If Index = 1 Then a$ = "#2 "
    strText = Trim$(txtTagLoc.Text)
    If Not IsNumber(strText) Then
        TagLocErr = "You must enter a number for Tag " & a$ & "Location"
    ElseIf CDbl(strText) < 0 Then
        TagLocErr = "Please enter a positive number for Tag " & a$ & "Location"
    End If
End Function

Private Function TagWeightErr(Index As Integer) As String
    'Returns an error msg if the indicated 'txtTagLoc()' control contains bad data
    Dim strText As String
    Dim a$
    
    If Index = 1 Then a$ = "#2 "
    strText = Trim$(txtTagWeight.Text)
    If Not IsNumber(strText) Then
        TagWeightErr = "You must enter a number for Tag " & a$ & "Weight"
    ElseIf CDbl(strText) < 0 Then
        TagWeightErr = "Please enter a positive number for Tag " & a$ & "Weight"
    End If
End Function

Private Function TagWtLimitErr(Index As Integer) As String
    'Returns an error msg if the indicated 'txtTagLoc()' control contains bad data
    Dim strText As String
    Dim a$
    
    If Index = 1 Then a$ = "#2 "
    strText = Trim$(txtTagWtLimit.Text)
    If Not IsNumber(strText) Then
        TagWtLimitErr = "You must enter a number for Tag " & a$ & "Weight-Limit"
    ElseIf CDbl(strText) < 0 Then
        TagWtLimitErr = "Please enter a positive number for Tag " & a$ & "Weight-Limit"
    End If
End Function

Private Sub cmdSave_Click()
    'Save changes to Active Truck and save as new chassis object file
    Dim sMsg As String
    Dim colComponents As New Collection
    Dim i%
    
    If Not UpdateActiveTruck Then
        Exit Sub
    End If
    
    For i% = 0 To lstComponents.ListCount - 1
        If lstComponents.Selected(i%) Then
            colComponents.Add CStr(i% + 1)  'Component Index
        End If
    Next i%
    
    sMsg = SaveChassisObject(cActiveTruck, colComponents, txtObjectDisplayName.Text, txtObjectFullName.Text)
    If sMsg <> "" Then
        MsgBox sMsg, vbCritical, "Failed to save Chassis"
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Save changes to Active Truck and exit
    If UpdateActiveTruck Then
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function UpdateActiveTruck() As Boolean
    'This sub will first verify everything on this form.  If no gross errors are found,
    'the data will put into the ActiveTruck object
    Dim Row As Integer
    Dim i%
    Dim strErr As String
    Dim strText As String
    Dim strCellData As String
    Dim cChassis As clsChassis
    Dim cTag As clsTag
    On Error GoTo errHandler
    
    UpdateActiveTruck = False 'default to fail
    
    'Verify-------------------------------------------------
    strErr = ""
    For Row = 0 To grdChassis.Rows - 1
        strText = grdChassis.Cell(flexcpText, Row, 1)
        strCellData = grdChassis.Cell(flexcpData, Row, 1)
        AppendError strErr, VerifyCell(strText, strCellData)
    Next Row
    
    For i% = 1 To cTags.Count
        With cTags(i%)
            If .Location = 0 Then
                AppendError strErr, "Tag " & i% & " Axle cannot be located at 0.0"
            End If
        End With
    Next i%
    
    If strErr <> "" Then
        strErr = "Please fix the following problems and try again:" & vbCrLf & strErr
        MsgBox strErr, vbExclamation, "Bad Data"
        Exit Function
    End If
    
    
    'Update-------------------------------------------------
    Set cChassis = New clsChassis
    'Fill in Chassis-level properties
    For Row = 0 To grdChassis.Rows - 1
        strText = Trim$(grdChassis.Cell(flexcpText, Row, 1))
        Select Case grdChassis.Cell(flexcpData, Row, 1)
        Case "FullName"
            cChassis.FullName = strText
        Case "DisplayName"
            cChassis.DisplayName = strText
        Case "WB"
            cChassis.WB = CDbl(strText) / cGlobalInfo.DistanceUnits.Multiplier
        Case "BackOfCab"
            cChassis.BackOfCab = CDbl(strText) / cGlobalInfo.DistanceUnits.Multiplier
        Case "WtFront"
            cChassis.WtFront = CDbl(strText) / cGlobalInfo.MassUnits.Multiplier
        Case "WtRear"
            cChassis.WtRear = CDbl(strText) / cGlobalInfo.MassUnits.Multiplier
        Case "WtLimitFront"
            cChassis.WtLimitFront = CDbl(strText) / cGlobalInfo.MassUnits.Multiplier
        Case "WtLimitRear"
            cChassis.WtLimitRear = CDbl(strText) / cGlobalInfo.MassUnits.Multiplier
        Case "WtLimitTotal"
            If CDbl(strText) = 0 Then
                cChassis.WtLimitTotal = cChassis.WtLimitFront + cChassis.WtLimitRear
            Else
                cChassis.WtLimitTotal = CDbl(strText) / cGlobalInfo.MassUnits.Multiplier
            End If
        Case "TandemSpacing"
            cChassis.TandemSpacing = CDbl(strText) / cGlobalInfo.DistanceUnits.Multiplier
        Case "PlacementAllowable"
            'user has no choice here, so enforce it
            cChassis.PlacementAllowable = paCenter
            grdChassis.Cell(flexcpText, Row, 1) = cChassis.PlacementAllowableString
        Case "StreetSideStd"
            cChassis.StreetSideStd = strText
        Case "CurbSideStd"
            cChassis.CurbSideStd = strText
        Case "WheelDia"
            cChassis.WheelDia = CDbl(strText) / cGlobalInfo.DistanceUnits.Multiplier
        Case "WheelY"
            cChassis.WheelY = CDbl(strText) / cGlobalInfo.DistanceUnits.Multiplier
        Case "FrameRailHt"
            cChassis.FrameRailHt = CDbl(strText) / cGlobalInfo.DistanceUnits.Multiplier
        Case "TwinSteerSeparation"
            cChassis.TwinSteerSeparation = CDbl(strText) / cGlobalInfo.DistanceUnits.Multiplier
        End Select
    Next Row
    
    'Program considers WB to be relative to center of TwinSteer
    cChassis.WB = cChassis.WB - cChassis.TwinSteerSeparation / 2
    
    'Correct Tag location to be relative to center of TwinSteer
    For i% = 1 To cTags.Count
        cTags(i%).Location = cTags(i%).Location - cChassis.TwinSteerSeparation / 2
    Next i%
    
    'Add Tag(s) info
    Set cChassis.Tags = New clsTags
    
    Set cChassis.Tags = cTags
    
    Set cActiveTruck.Chassis = Nothing
    Set cActiveTruck.Chassis = cChassis
    UpdateActiveTruck = True
CleanUp:
    Set cChassis = Nothing
    Set cTag = Nothing
    Exit Function

errHandler:
    ErrorIn "frmEditChassis.UpdateActiveTruck"
    Resume CleanUp
End Function

Private Function VerifyCell(strText As String, strCellData As String) As String
    'Checks a grid cell for valid data
    VerifyCell = ""
    Select Case strCellData
    Case "FullName"
        'nothing to check, really.
    Case "DisplayName"
        'nothing to check, really.
    Case "WB"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'WB'"
        ElseIf CDbl(strText) <= 0 Then
            VerifyCell = "Please enter a positive, non-zero number for 'WB'"
        End If
    Case "BackOfCab"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'BackOfCab'"
        End If
    Case "WtFront"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'WtFront'"
        ElseIf CDbl(strText) <= 0 Then
            VerifyCell = "Please enter a positive, non-zero number for 'WtFront'" '
        End If
    Case "WtRear"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'WtRear'"
        ElseIf CDbl(strText) <= 0 Then
            VerifyCell = "Please enter a positive, non-zero number for 'WtRear'"
        End If
    Case "WtLimitFront"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'WtLimitFront'"
        ElseIf CDbl(strText) < 0 Then
            VerifyCell = "Please enter a positive number for 'WtLimitFront'"
        End If
    Case "WtLimitRear"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'WtLimitRear'"
        ElseIf CDbl(strText) < 0 Then
            VerifyCell = "Please enter a positive number for 'WtLimitRear'"
        End If
    Case "WtLimitTotal"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'WtLimitTotal'"
        ElseIf CDbl(strText) < 0 Then
            VerifyCell = "Please enter a positive number for 'WtLimitTotal'"
        End If
    Case "TandemSpacing"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'TandemSpacing'"
        ElseIf CDbl(strText) < 0 Then
            VerifyCell = "Please enter a positive number for 'TandemSpacing'"
        End If
    Case "PlacementAllowable"
        'user has no choice here, so enforce it
        If UCase$(strText) <> "CENTER" Then
            VerifyCell = "'PlacementAllowable' must always be set to 'Center' for a chassis."
        End If
    Case "WheelDia"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'WheelDia'"
        ElseIf CDbl(strText) < 0 Then
            VerifyCell = "Please enter a positive number for 'WheelDia'"
        End If
    Case "WheelY"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a negative number for 'WheelY'"
        ElseIf CDbl(strText) > 0 Then
            VerifyCell = "You must enter a negative number for 'WheelY'"
        End If
    Case "FrameRailHt"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'FrameRailHt'"
        ElseIf CDbl(strText) < 0 Then
            VerifyCell = "Please enter a positive number for 'FrameRailHt'"
        End If
    Case "TwinSteerSeparation"
        If Not IsNumber(strText) Then
            VerifyCell = "You must enter a number for 'TwinSteerSeparation'"
        ElseIf CDbl(strText) < 0 Then
            VerifyCell = "Please enter a positive number for 'TwinSteerSeparation'"
        End If
    End Select
End Function

Private Function AppendError(ByRef strErr As String, strAdd As String)
    If strAdd <> "" Then
        'Add this string
        strErr = strErr & vbCrLf & strAdd
    End If
End Function

Private Function InitTags(tabTags As C1Tab, Optional bRefresh As Boolean = False)
    Dim i%
    Dim dblTemp As Double
    Dim NumTags As Integer
    Dim cTag As New clsTag

    
    If bRefresh Then
        NumTags = cTags.Count
    Else
        Set cTags = New clsTags
        NumTags = cActiveTruck.Chassis.Tags.Count
    End If
    tabTags.Caption = ""
    tabTags.RemoveTab 0
    LastTagIndex = -1
    If NumTags = 0 Then 'No tags to display
        txtTagLoc.Enabled = False
        txtTagWeight.Enabled = False
        txtTagWtLimit.Enabled = False
        txtTagLoc.Text = ""
        txtTagWeight.Text = ""
        txtTagWtLimit.Text = ""
        Exit Function
    Else 'One or more tags
        txtTagLoc.Enabled = True
        txtTagWeight.Enabled = True
        txtTagWtLimit.Enabled = True
    End If
    
    'Add tags to display and (if necessary) a temporary class collection
    For i% = 1 To NumTags
        tabTags.AddTab "Tag " & i%
        If Not bRefresh Then
            'Add tag to temp collection and fill tag with data
            Set cTag = New clsTag
            With cTag
                .DownwardForce = cActiveTruck.Chassis.Tags(i%).DownwardForce
                .ForceToPressure = cActiveTruck.Chassis.Tags(i%).ForceToPressure
                .Location = cActiveTruck.Chassis.Tags(i%).Location + cActiveTruck.Chassis.TwinSteerSeparation / 2
                .Weight = cActiveTruck.Chassis.Tags(i%).Weight
                .WtLimit = cActiveTruck.Chassis.Tags(i%).WtLimit
            End With
            cTags.Add cTag
        End If
    Next
        
    'Set the first tab as the active tab
    tabTags.CurrTab = NumTags - 1
    tabTags.FirstTab = 0
    'Tag Location
    dblTemp = cTags(NumTags).Location * cGlobalInfo.DistanceUnits.Multiplier
    txtTagLoc.Text = Format(dblTemp, sLenDispFormat)
    'Tag Wt
    dblTemp = cTags(NumTags).Weight * cGlobalInfo.MassUnits.Multiplier
    txtTagWeight.Text = Format(dblTemp, sLenDispFormat)
    'Tag Load Limit
    dblTemp = cTags(NumTags).WtLimit * cGlobalInfo.MassUnits.Multiplier
    txtTagWtLimit.Text = Format(dblTemp, sLenDispFormat)

    Set cTag = Nothing
End Function

Private Function AddTag(tabTags As C1Tab)
    Dim cTag As New clsTag
    
    If TypeName(cTags) <> "Nothing" Then
        If cTags.Count >= 4 Then
            MsgBox "You cannot have more than 4 Tag axles", vbExclamation, "Invalid Operation"
            Exit Function
        End If
    End If
    
    cTags.Add cTag
    If cTags.Count = 1 Then
        txtTagLoc.Enabled = True
        txtTagWeight.Enabled = True
        txtTagWtLimit.Enabled = True
    End If
    txtTagLoc.Text = "0"
    txtTagWeight.Text = "0"
    txtTagWtLimit.Text = "0"
    tabTags.AddTab "Tag " & cTags.Count
    tabTagDisplay.CurrTab = cTags.Count - 1
    Set cTag = Nothing
End Function

Private Sub tabTagDisplay_Click()
    Dim dblTemp As Double
    
    If tabTagDisplay.CurrTab = LastTagIndex Then Exit Sub

    With cTags(tabTagDisplay.CurrTab + 1)
        'Tag Location
        dblTemp = .Location * cGlobalInfo.DistanceUnits.Multiplier
        txtTagLoc.Text = Format(dblTemp, sLenDispFormat)
        'Tag Wt
        dblTemp = .Weight * cGlobalInfo.MassUnits.Multiplier
        txtTagWeight.Text = Format(dblTemp, sLenDispFormat)
        'Tag Load Limit
        dblTemp = .WtLimit * cGlobalInfo.MassUnits.Multiplier
        txtTagWtLimit.Text = Format(dblTemp, sLenDispFormat)
    End With
    LastTagIndex = tabTagDisplay.CurrTab
End Sub


Private Sub picDelete_Click()
    If tabTagDisplay.NumTabs = 0 Then Exit Sub
    
    cTags.Remove tabTagDisplay.CurrTab + 1
    InitTags tabTagDisplay, True
End Sub

Private Sub picNew_Click()
    AddTag tabTagDisplay
End Sub

