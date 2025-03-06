VERSION 5.00
Object = "{D3F92121-EFAA-4B5C-B91B-3D6A8FFD1477}#1.0#0"; "vsdraw8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmMonitor 
   Caption         =   "Monitor Progress"
   ClientHeight    =   8496
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12096
   Icon            =   "frmMonitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8496
   ScaleWidth      =   12096
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8496
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12096
      _cx             =   21336
      _cy             =   14986
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
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic szrEmulsion 
         Height          =   480
         Left            =   132
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   252
         Width           =   936
         _cx             =   1651
         _cy             =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "Emulsion"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   2
         ChildSpacing    =   2
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   3
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   1
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   -1  'True
         GridRows        =   0
         GridCols        =   0
         Frame           =   4
         FrameStyle      =   3
         FrameWidth      =   1
         FrameColor      =   -2147483630
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Label lblEmulPct 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "33%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   228
            Left            =   192
            TabIndex        =   25
            Top             =   240
            Width           =   552
         End
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "ñ"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "ò"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   3720
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "ñ"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   4080
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "ñ"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   4440
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "ñ"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   4
         Left            =   4800
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "ò"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   4080
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "ò"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   4440
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "ò"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   16.2
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   4
         Left            =   4800
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox txtWarnings 
         Height          =   912
         Left            =   3720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2880
         Width           =   8292
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   612
         Left            =   11040
         TabIndex        =   5
         Top             =   7800
         Width           =   972
      End
      Begin VB.Timer tmrButton 
         Interval        =   100
         Left            =   7200
         Top             =   2520
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Preview"
         Height          =   492
         Left            =   3720
         TabIndex        =   4
         Top             =   7920
         Width           =   1452
      End
      Begin VB.CheckBox chkOnRoad 
         Caption         =   "Find loading that is road-legal"
         Height          =   252
         Left            =   8160
         TabIndex        =   3
         Top             =   2400
         Width           =   3372
      End
      Begin VB.Frame fraTag 
         Caption         =   "Tag Axles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2052
         Left            =   8160
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   3732
         Begin C1SizerLibCtl.C1Tab tabTagDisplay 
            Height          =   1452
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   3252
            _cx             =   5736
            _cy             =   2561
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
               Height          =   1164
               Left            =   12
               ScaleHeight     =   1164
               ScaleWidth      =   3228
               TabIndex        =   27
               Top             =   276
               Width           =   3228
               Begin VB.TextBox txtTag 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   28
                  Text            =   "0"
                  Top             =   360
                  Width           =   972
               End
               Begin VB.Label lblMassUnits 
                  BackStyle       =   0  'Transparent
                  Caption         =   "kg"
                  Height          =   276
                  Index           =   1
                  Left            =   1920
                  TabIndex        =   35
                  Top             =   840
                  Width           =   492
               End
               Begin VB.Label lblForce 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Applied Force"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   276
                  Left            =   120
                  TabIndex        =   34
                  Top             =   120
                  Width           =   1332
               End
               Begin VB.Label lblMassUnits 
                  BackStyle       =   0  'Transparent
                  Caption         =   "kg"
                  Height          =   276
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   33
                  Top             =   410
                  Width           =   492
               End
               Begin VB.Label lblTagLimit 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Allowable:"
                  Height          =   252
                  Index           =   0
                  Left            =   120
                  TabIndex        =   32
                  Top             =   840
                  Width           =   972
               End
               Begin VB.Label lblTagLimit 
                  BackStyle       =   0  'Transparent
                  Caption         =   "64599"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   276
                  Index           =   1
                  Left            =   1200
                  TabIndex        =   31
                  Top             =   840
                  Width           =   732
               End
               Begin VB.Label lblPSI 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Air Pres."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   276
                  Index           =   0
                  Left            =   2040
                  TabIndex        =   30
                  Top             =   120
                  Width           =   972
               End
               Begin VB.Label lblPSI 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Air Pres."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   276
                  Index           =   1
                  Left            =   2040
                  TabIndex        =   29
                  ToolTipText     =   "Provides estimate for tag air pressure to achieve indicated ""Applied Force"""
                  Top             =   408
                  Width           =   972
               End
            End
         End
      End
      Begin VB.CommandButton cmdManual 
         Caption         =   "Manually Edit"
         Height          =   372
         Left            =   10320
         TabIndex        =   1
         Top             =   240
         Width           =   1692
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   5532
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   3372
         _ExtentX        =   5948
         _ExtentY        =   9758
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VSDraw8LibCtl.VSDraw vsdBody 
         Height          =   2172
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7812
         _cx             =   13779
         _cy             =   3831
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScaleLeft       =   0
         ScaleTop        =   0
         ScaleHeight     =   1000
         ScaleWidth      =   1000
         PenColor        =   0
         PenWidth        =   0
         PenStyle        =   0
         BrushColor      =   -2147483633
         BrushStyle      =   0
         TextColor       =   -2147483640
         TextAngle       =   0
         TextAlign       =   0
         BackStyle       =   0
         LineSpacing     =   100
         EmptyColor      =   -2147483636
         PageWidth       =   0
         PageHeight      =   0
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   30
         SmallChangeVert =   30
         Track           =   -1  'True
         MouseScroll     =   -1  'True
         ProportionalBars=   -1  'True
         Zoom            =   100
         ZoomMode        =   0
         KeepTextAspect  =   -1  'True
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
      Begin MSFlexGridLib.MSFlexGrid msgTanks 
         Height          =   2412
         Left            =   3720
         TabIndex        =   17
         Top             =   5400
         Width           =   8292
         _ExtentX        =   14626
         _ExtentY        =   4255
         _Version        =   393216
         Rows            =   3
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483626
         BackColorBkg    =   -2147483633
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid msgLoading 
         Height          =   950
         Left            =   3720
         TabIndex        =   18
         Top             =   4200
         Width           =   8292
         _ExtentX        =   14626
         _ExtentY        =   1672
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   -2147483626
         BackColorBkg    =   -2147483633
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Info"
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
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   3372
      End
      Begin VB.Label lblWarnings 
         BackStyle       =   0  'Transparent
         Caption         =   "Warnings"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   2640
         Width           =   5772
      End
      Begin VB.Label lblTankFill 
         BackStyle       =   0  'Transparent
         Caption         =   "Tank Fill"
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
         Left            =   3720
         TabIndex        =   21
         Top             =   5160
         Width           =   2172
      End
      Begin VB.Label lblAxleLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Axle Loading"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   3960
         Width           =   2172
      End
      Begin VB.Label lblProductTanks 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Tank Level Diagram"
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
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   3372
      End
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cCurTruck As clsTruck 'used in this form as a link to the active truck
Private bMouseBtnHeld As Boolean
Private bDirectionUp As Boolean 'TRUE= Up Button held, FALSE= Down Button held
Private intBtnIndex As Integer
Private lCounts As Long 'increases as long as up/down button is held

Public Sub MonitorDwg(cTruck As clsTruck, DisplayText As String)
    On Error GoTo errHandler
    
    Me.Show
    
    C1Elastic1.AutoSizeChildren = azNone
    szrEmulsion.Visible = False
    chkOnRoad.Visible = False
    fraTag.Visible = False
    HideButtons
    Set cCurTruck = cTruck
    UpdateDrawing vsdBody, cTruck
    txtWarnings.Text = DisplayText
    UpdateInfoBox
    UpdateLoading
    UpdateTanks
    cmdManual.Visible = True
    C1Elastic1.AutoSizeChildren = azProportional
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.MonitorDwg(cTruck,DisplayText)", Array(cTruck, DisplayText)
End Sub

Public Sub ManualFill(cTruck As clsTruck, Optional bEmpty As Boolean = True)
    'Called from elsewhere to allow manually editing tank volumes
    Dim sErr As String
    Dim MinWtFront As Double
    Dim MinWtRear As Double
    On Error GoTo errHandler
    
    Set cCurTruck = cTruck
    
    If bEmpty Then EmptyTruck
    
    Me.Show
    
    C1Elastic1.AutoSizeChildren = azNone
    szrEmulsion.Visible = IsBlendConfiguration(cTruck)
    cmdManual.Visible = False
    chkOnRoad.Value = Abs(CInt(bIsOnRoad))
    InitForm
    chkOnRoad.Visible = True
    
    sErr = sErr & vbCrLf & LoadingViolations(cCurTruck, bIsOnRoad)
    
    UpdateDrawing vsdBody, cCurTruck
    txtWarnings.Text = sErr
    UpdateInfoBox
    
    ShowButtons
    C1Elastic1.AutoSizeChildren = azProportional

    Exit Sub
errHandler:
    ErrorIn "frmMonitor.ManualFill(cTruck,bEmpty)", Array(cTruck, bEmpty)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EmptyTruck
End Sub

Private Function UpdateInfoBox() As Boolean
    'This sub fills the Info box with all sorts of loading information
    Dim itmX As ListItem
    Dim clmX As ColumnHeader
    Dim a$, b$
    Dim i%
    Dim dblLoad As Double
    Dim dblGVW As Double
    Dim dblTemp As Double
    Dim dblFuelVol As Double
    Dim dblFuelMass As Double
    Dim cEmulTanks As New clsTanks
    Dim cANTanks As New clsTanks
    On Error GoTo errHandler
    
    Set cANTanks = ANTanks(cCurTruck)
    Set cEmulTanks = EmulsionTanks(cCurTruck)
        
    lvwInfo.ListItems.Clear
    lvwInfo.ColumnHeaders.Clear
    Set clmX = lvwInfo.ColumnHeaders.Add(, , "Item", 600)
    Set clmX = lvwInfo.ColumnHeaders.Add(, , "Value")
    
    'Total AN (if applicable)
    If cANTanks.Count > 0 Then
        'Weight Total
        dblTemp = cCurTruck.Body.MassANTotal * cGlobalInfo.MassUnits.Multiplier
        a$ = "AN Wt Total"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        Set itmX = lvwInfo.ListItems.Add(, , a$)
            itmX.SubItems(1) = b$
        'Volume Total
        dblTemp = cCurTruck.Body.MassANTotal / cGlobalInfo.DensityAN * cGlobalInfo.VolumeUnits.Multiplier
        a$ = "AN Vol Total"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
        Set itmX = lvwInfo.ListItems.Add(, , a$)
            itmX.SubItems(1) = b$
        'Tank Levels
        For i% = 1 To cANTanks.Count
            a$ = cANTanks(i%).DisplayName
            dblTemp = cANTanks(i%).CurVol * cGlobalInfo.VolumeUnits.Multiplier
            b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
            Set itmX = lvwInfo.ListItems.Add(, , a$)
                itmX.SubItems(1) = b$
        Next
    End If
    
    'Total Emulsion (if applicable)
    If cEmulTanks.Count > 0 Then
        Set itmX = lvwInfo.ListItems.Add(, , " ") 'Blank Line
            itmX.SubItems(1) = " "
        'Percent Emulsion (if applicable)
        If cEmulTanks.Count > 0 And cANTanks.Count > 0 Then
            a$ = "%Emul"
            If (cCurTruck.Body.MassANTotal / (1# - dblDFOPct) + cCurTruck.Body.MassEmulTotal) = 0 Then
                'Division by zero
                b$ = "0%"
            Else
                dblTemp = 100# * cCurTruck.Body.MassEmulTotal / (cCurTruck.Body.MassANTotal / (1# - dblDFOPct) + cCurTruck.Body.MassEmulTotal)
                b$ = Format(dblTemp, "#,##0") & "%"
            End If
            Set itmX = lvwInfo.ListItems.Add(, , a$)
                itmX.SubItems(1) = b$
            lblEmulPct.Caption = b$
        End If
        'Weight Total
        dblTemp = cCurTruck.Body.MassEmulTotal * cGlobalInfo.MassUnits.Multiplier
        a$ = "Emul Wt Total"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        Set itmX = lvwInfo.ListItems.Add(, , a$)
            itmX.SubItems(1) = b$
        'Volume Total
        dblTemp = cCurTruck.Body.MassEmulTotal / cGlobalInfo.DensityEmul * cGlobalInfo.VolumeUnits.Multiplier
        a$ = "Emul Vol Total"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
        Set itmX = lvwInfo.ListItems.Add(, , a$)
            itmX.SubItems(1) = b$
        'Tank Levels
        For i% = 1 To cEmulTanks.Count
            a$ = cEmulTanks(i%).DisplayName
            dblTemp = cEmulTanks(i%).CurVol * cGlobalInfo.VolumeUnits.Multiplier
            b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
            Set itmX = lvwInfo.ListItems.Add(, , a$)
                itmX.SubItems(1) = b$
        Next
    End If
    
    'Total Fuel (if applicable)
    dblFuelMass = 0#
    dblFuelVol = 0#
    For i% = 1 To cCurTruck.Components.Count
        Select Case cCurTruck.Components(i%).ContentsType
        Case ctFuel
            If cCurTruck.Components(i%).Capacity.DefaultWtContents = "" Then
                dblTemp = cCurTruck.Components(i%).Capacity.CurVol * cGlobalInfo.VolumeUnits.Multiplier
                dblFuelVol = dblFuelVol + dblTemp
                dblTemp = cCurTruck.Components(i%).Capacity.CurVol * cCurTruck.Components(i%).Capacity.DensityContents * cGlobalInfo.MassUnits.Multiplier
                dblFuelMass = dblFuelMass + dblTemp
            End If
        End Select
    Next
    If dblFuelVol > 0# Then
        'There is some fuel...
        Set itmX = lvwInfo.ListItems.Add(, , " ") 'Blank Line
            itmX.SubItems(1) = " "
        'Show Wt
        a$ = "Fuel Wt Total"
        b$ = Format(dblFuelMass, "#,##0 ") & cGlobalInfo.MassUnits.Display
        Set itmX = lvwInfo.ListItems.Add(, , a$)
            itmX.SubItems(1) = b$
        'Show Volume
        a$ = "Fuel Vol Total"
        b$ = Format(dblFuelVol, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
        Set itmX = lvwInfo.ListItems.Add(, , a$)
            itmX.SubItems(1) = b$
    End If
    
    
    'Total Product (if applicable)
    If Not (cEmulTanks.Count = 0 Or cANTanks.Count = 0) Then
        Set itmX = lvwInfo.ListItems.Add(, , " ") 'Blank Line
            itmX.SubItems(1) = " "
        dblTemp = (cCurTruck.Body.MassANTotal + cCurTruck.Body.MassEmulTotal) * cGlobalInfo.MassUnits.Multiplier
        dblTemp = dblTemp + dblFuelMass
        a$ = "Total Product"
            b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        Set itmX = lvwInfo.ListItems.Add(, , a$)
            itmX.SubItems(1) = b$
        Set itmX = lvwInfo.ListItems.Add(, , " ") 'Blank Line
            itmX.SubItems(1) = " "
    End If
        
    'Levels of components (water, gassing, etc.)
    For i% = 1 To cCurTruck.Components.Count
        Select Case cCurTruck.Components(i%).ContentsType
        Case ctNone, ctFuel
            'Don't show
            a$ = ""
            b$ = ""
        Case Else
            a$ = cCurTruck.Components(i%).DisplayName
            If cCurTruck.Components(i%).Capacity.DefaultWtContents = "" Then
                dblTemp = cCurTruck.Components(i%).Capacity.CurVol * cGlobalInfo.VolumeUnits.Multiplier
                b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
            Else
                b$ = FormattedLevel(cCurTruck.Components(i%))
            End If
        End Select
        If a$ <> "" Then
            Set itmX = lvwInfo.ListItems.Add(, , a$)
                itmX.SubItems(1) = b$
        End If
    Next
    
    'Manufacturer's Max Load Spec for Front axle/tire
    Set itmX = lvwInfo.ListItems.Add(, , " ") 'Blank Line
        itmX.SubItems(1) = " "
    dblLoad = cCurTruck.Chassis.WtLimitFront * cGlobalInfo.MassUnits.Multiplier
    a$ = "MfgLimit Front"
    b$ = Format(dblLoad, "#,##0 ") & cGlobalInfo.MassUnits.Display
    Set itmX = lvwInfo.ListItems.Add(, , a$)
        itmX.SubItems(1) = b$
    'Manufacturer's Max Load Spec for Rear axle/tire
    dblLoad = cCurTruck.Chassis.WtLimitRear * cGlobalInfo.MassUnits.Multiplier
    a$ = "MfgLimit Rear"
    b$ = Format(dblLoad, "#,##0 ") & cGlobalInfo.MassUnits.Display
    Set itmX = lvwInfo.ListItems.Add(, , a$)
        itmX.SubItems(1) = b$
    
'    'Product VCG
'    a$ = "Product VCoG"
'    dblTemp = ProductVCOG(cCurTruck.Body) * cGlobalInfo.DistanceUnits.Multiplier
'    b$ = Format(dblTemp, "0.# ") & cGlobalInfo.DistanceUnits.Display
'    Set itmX = lvwInfo.ListItems.Add(, , a$)
'        itmX.SubItems(1) = b$
    
    LV_AutoSizeColumn lvwInfo
    
    Set cANTanks = Nothing
    Set cEmulTanks = Nothing
    Set itmX = Nothing
    Exit Function
errHandler:
    ErrorIn "frmMonitor.UpdateInfoBox"
End Function

Private Sub cmdManual_Click()
    ManualFill cCurTruck, False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub HideButtons()
    Dim i%
    
    For i% = 1 To 4
        cmdUp(i%).Visible = False
        cmdDown(i%).Visible = False
    Next i%
End Sub

Private Sub ShowButtons()
    Dim i%
    Dim colCtrs As Collection
    On Error GoTo errHandler
    
    Set colCtrs = ButtonCenters
    HideButtons 'hide all first
    
    For i% = 1 To colCtrs.Count
        
        cmdUp(i%).Left = colCtrs.Item(CStr(i%)) * vsdBody.Width + vsdBody.Left _
                         - cmdUp(i%).Width / 2
        cmdDown(i%).Left = cmdUp(i%).Left
        
        cmdUp(i%).Visible = True
        cmdDown(i%).Visible = True
    Next i%
    Set colCtrs = Nothing
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.ShowButtons"
End Sub

Private Sub Increase(Index As Integer, Optional bFill As Boolean = False)
    'Increase volume in Tank(Index) if possible
    Dim VolIncr As Double
    On Error GoTo errHandler
    
    VolIncr = cCurTruck.Body.Tanks(Index).Volume / 60
    
    
    If cCurTruck.Body.Tanks(Index).CurVol + VolIncr > cCurTruck.Body.Tanks(Index).Volume _
       Or bFill Then
        'Can't go higher
        cCurTruck.Body.Tanks(Index).CurVol = cCurTruck.Body.Tanks(Index).Volume
    Else
        cCurTruck.Body.Tanks(Index).CurVol = cCurTruck.Body.Tanks(Index).CurVol + VolIncr
    End If
      
    RecalcAndUpdate
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.Increase(index,bFill)", Array(Index, bFill)
End Sub

Private Sub Decrease(Index As Integer)
    'Decrease volume in Tank(Index) if possible
    Dim VolIncr As Double
    On Error GoTo errHandler
    
    VolIncr = cCurTruck.Body.Tanks(Index).Volume / 60
    
    
    If (cCurTruck.Body.Tanks(Index).CurVol - VolIncr) < 0 Then
        'Can't go lower
        cCurTruck.Body.Tanks(Index).CurVol = 0#
    Else
        cCurTruck.Body.Tanks(Index).CurVol = cCurTruck.Body.Tanks(Index).CurVol - VolIncr
    End If
        
    RecalcAndUpdate
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.Decrease(index)", Index
End Sub

Private Sub RecalcAndUpdate()
    Dim sErr As String
    On Error GoTo errHandler
    
    sErr = ReCalc(cCurTruck)
    
    sErr = sErr & vbCrLf & LoadingViolations(cCurTruck, bIsOnRoad)
    
    UpdateDrawing vsdBody, cCurTruck
    txtWarnings.Text = sErr
    UpdateInfoBox
    UpdateLoading
    UpdateTanks
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.RecalcAndUpdate"
End Sub


Private Sub txtTag_Change()
    'Force validation of any changes
    Dim bDummy As Boolean
    
    Call txtTag_Validate(bDummy)
End Sub

Private Sub txtTag_Validate(Cancel As Boolean)
    'Check Tag value.  If OK, convert to internal units and save value
    Dim dblNum As Double
    Dim dblMax As Double 'Max value (in approp. eng. units)
    Dim sLimit As String
    Dim i%
    On Error GoTo errHandler
    
    If Trim$(txtTag.Text) = "" Then
        'will treat this as zero
    ElseIf Not IsNumber(txtTag.Text) Then
        Cancel = True
        MsgBox "Please enter a valid number for Tag Force"
        Exit Sub
    End If
    
    i% = tabTagDisplay.CurrTab + 1
    'dblMax = TagLimit(cActiveTruck, i%) * cGlobalInfo.MassUnits.Multiplier
    dblMax = cActiveTruck.Chassis.Tags(i%).WtLimit * cGlobalInfo.MassUnits.Multiplier
        
    If Trim$(txtTag.Text) = "" Then
        dblNum = 0
    Else
        dblNum = Var2Dbl(txtTag.Text)
    End If
    
    If dblNum < 0 Or (dblNum > dblMax + 0.5) Then
        sLimit = Format(dblMax, "###") & " (" & cGlobalInfo.MassUnits.Display & ")"
        MsgBox "Please enter a number from 0 to " & sLimit & " for Tag " & i%
        Cancel = True
        Exit Sub
    End If
    cActiveTruck.Chassis.Tags(i%).DownwardForce = dblNum / cGlobalInfo.MassUnits.Multiplier
    vsdBody.Tag = 0 'this will force a complete redraw
    UpdateTagDisplay
    RecalcAndUpdate
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.txtTag_Validate(Cancel)", Cancel
End Sub


Private Sub EmptyTruck()
    Dim i%
    Dim sDummy As String
    On Error GoTo errHandler
    
    For i% = 1 To cCurTruck.Body.Tanks.Count
        cCurTruck.Body.Tanks(i%).CurVol = 0#
    Next
    sDummy = ReCalc(cCurTruck)
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.EmptyTruck"
End Sub

Private Sub InitForm()
    Dim dblVal As Double
    Dim i%
    On Error GoTo errHandler
    
    'Tag options
    If cActiveTruck.Chassis.Tags.Count > 0 Then
        'One or more tags
        fraTag.Visible = True
        tabTagDisplay.Caption = ""
        tabTagDisplay.RemoveTab 0
        'Add tags to display and (if necessary) a temporary class collection
        For i% = 1 To cActiveTruck.Chassis.Tags.Count
            If cActiveTruck.Chassis.Tags(i%).Location < cActiveTruck.Chassis.WB Then
                tabTagDisplay.AddTab i% & " (Pusher)"
            Else
                tabTagDisplay.AddTab i% & " (Tag)"
            End If
        Next
        'Set the first tab as the active tab
        tabTagDisplay.CurrTab = cActiveTruck.Chassis.Tags.Count - 1
        tabTagDisplay.FirstTab = 0
        
        UpdateTagDisplay
    Else 'No Tags
        fraTag.Visible = False
    End If
    
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.InitForm"
End Sub

Private Sub chkOnRoad_Click()
    On Error GoTo errHandler
    
    If chkOnRoad.Value = 1 Then
        bIsOnRoad = True
    Else
        bIsOnRoad = False
    End If
    UpdateTagDisplay
    RecalcAndUpdate
    frmMain.chkOnRoad.Value = chkOnRoad.Value
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.chkOnRoad_Click"
End Sub

Private Sub UpdateTagDisplay()
    Dim dblVal As Double
    Dim i%
    On Error GoTo errHandler
    
    'Tag options
    If cActiveTruck.Chassis.Tags.Count > 0 Then
        i% = tabTagDisplay.CurrTab + 1
        dblVal = cActiveTruck.Chassis.Tags(i%).DownwardForce * cGlobalInfo.MassUnits.Multiplier
        txtTag.Text = Format(dblVal, "#0")
        If cActiveTruck.Chassis.Tags(i%).ForceToPressure > 0 Then
            dblVal = cActiveTruck.Chassis.Tags(i%).DownwardForce * cActiveTruck.Chassis.Tags(i%).ForceToPressure
            lblPSI(1).Caption = Format(dblVal, "#,##0")
        Else
            lblPSI(1).Caption = "--"
        End If
        'dblVal = TagLimit(cActiveTruck, i%) * cGlobalInfo.MassUnits.Multiplier
        dblVal = cActiveTruck.Chassis.Tags(i%).WtLimit * cGlobalInfo.MassUnits.Multiplier
        lblTagLimit(1).Caption = Format(dblVal, "#,##0")
    End If
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.UpdateTagDisplay"
End Sub

Private Sub tabTagDisplay_Click()
    UpdateTagDisplay
End Sub


'------- UP Button -----------------------------------------------------------------
Private Sub cmdUp_Click(Index As Integer)
    Dim bCtrl As Boolean
    bCtrl = (CBool(GetKeyState(vbKeyControl) And &H8000))
    Increase Index, bCtrl 'handles cases where mouse not used
End Sub

Private Sub cmdUp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
        bMouseBtnHeld = True
        bDirectionUp = True
        intBtnIndex = Index
End Sub


Private Sub cmdUp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseBtnHeld = False
End Sub
'-----------------------------------------------------------------------------------

'-------  Down Button --------------------------------------------------------------
Private Sub cmdDown_Click(Index As Integer)
    Decrease Index
End Sub

Private Sub cmdDown_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
        bMouseBtnHeld = True
        bDirectionUp = False
        intBtnIndex = Index
End Sub

Private Sub cmdDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseBtnHeld = False
End Sub
'-----------------------------------------------------------------------------------

Private Sub tmrButton_Timer()
    'Increase timer as long as up/down button is pressed
    On Error GoTo errHandler
    
    If bMouseBtnHeld Then
        lCounts = lCounts + 1
        If lCounts > 3 Then
            If bDirectionUp Then
                Increase intBtnIndex
            Else
                Decrease intBtnIndex
            End If
        End If
    Else
        lCounts = 0
    End If
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.tmrButton_Timer"
End Sub

Private Sub cmdPrint_Click()
    'Show print-preview window
    frmReport.PrintPreview cCurTruck
End Sub

Private Sub UpdateLoading()
    Dim cAxles As clsAxleGroups
    Dim dblVal As Double
    Dim i%
    On Error GoTo errHandler
    
    'Now spit out the headings at the indicated spacing
    Set cAxles = LoadingSummary(cCurTruck, bIsOnRoad)
    
    msgLoading.Cols = cAxles.Count + 1
    
    For i% = 1 To cAxles.Count
        With msgLoading
            .Redraw = False
            .Row = 0
            .Col = i%
            .CellFontBold = True
            .Text = cAxles(i%).sDescription & " (" & cGlobalInfo.MassUnits.Display & ")"
            .CellAlignment = flexAlignCenterCenter
            
            .Row = 1
            .Col = i%
            dblVal = cAxles(i%).AllowableLd * cGlobalInfo.MassUnits.Multiplier
            .CellFontBold = False
            .Text = Format(dblVal, "#,##0")
            .CellAlignment = flexAlignCenterCenter
        
            .Row = 2
            .Col = i%
            .CellFontBold = False
            dblVal = cAxles(i%).ActualLd * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblVal, "#,##0")
            .CellAlignment = flexAlignCenterCenter
        End With
    Next i%
    
    With msgLoading
        .Row = 1
        .Col = 0
        .CellFontBold = True
        .Text = "Limits"
        
        .Row = 2
        .Col = 0
        .CellFontBold = True
        .Text = "Calculated Loading"
        .ColAlignment(0) = flexAlignRightCenter
        .Redraw = True
    End With
    SetGridColumnWidth msgLoading
    Set cAxles = Nothing
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.UpdateLoading"
End Sub

Private Sub UpdateTanks()
    Dim cAxles As clsAxleGroups
    Dim dblVal As Double
    Dim i%
    Dim a$
    Dim dblTotalProduct As Double
    Dim dblTemp As Double
    Dim bNotIntialized As Boolean
    Dim intRow As Integer
    On Error GoTo errHandler
    
    With msgTanks
        .Redraw = False
        If .Rows < 5 Then
            'This must be first call, so the table is not initialized yet
            .Clear
            .Cols = 6
            .Rows = 2
            bNotIntialized = True
        End If
    
        'Headings ==============================================
        .Row = 0
        .Col = 0
        .ColAlignment(0) = flexAlignRightCenter
        .CellFontBold = True
        .Text = "Tank"
        
        .Col = 1
        .ColAlignment(1) = flexAlignRightCenter
        .CellFontBold = True
        .Text = "Contents"
        
        .Col = 2
        .ColAlignment(2) = flexAlignRightCenter
        .CellFontBold = True
        .Text = "Density"
        
        .Col = 3
        .ColAlignment(3) = flexAlignRightCenter
        .CellFontBold = True
        .Text = "Contents (" & cGlobalInfo.MassUnits.Display & ")"
        
        .Col = 4
        .ColAlignment(4) = flexAlignRightCenter
        .CellFontBold = True
        .Text = "%Full"
        
        .Col = 5
        .ColAlignment(5) = flexAlignLeftCenter
        .CellFontBold = True
        .Text = "Fill"
    
        'Tanks =================================================
        intRow = 1
        .Row = intRow
        .CellFontBold = False
        For i% = 1 To cCurTruck.Body.Tanks.Count
            'Tank
            .Col = 0
            a$ = Chr$(64 + i%)
            .Text = a$
            'Contents
            .Col = 1
            If cCurTruck.Body.Tanks(i%).CurTankUse = ttAN Then
                .Text = "AN Prill"
            Else
                .Text = "Emulsion"
            End If
            'Density
            .Col = 2
            .Text = Format(cCurTruck.Body.Tanks(i%).DensityContents, "0.##")
            'Contents
            .Col = 3
            dblTemp = cCurTruck.Body.Tanks(i%).CurVol * cCurTruck.Body.Tanks(i%).DensityContents
            dblTotalProduct = dblTotalProduct + dblTemp
            dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblTemp, "#,##0 ") '& cGlobalInfo.MassUnits.Display
            '%Full
            .Col = 4
            dblTemp = cCurTruck.Body.Tanks(i%).CurVol / cCurTruck.Body.Tanks(i%).Volume * 100#
            .Text = Format(dblTemp, "#0") & "%"
            'Fill
            .Col = 5
            dblTemp = cCurTruck.Body.Tanks(i%).CurStkHt * cGlobalInfo.DistanceUnits.Multiplier
            If cGlobalInfo.DistanceUnits.Display = "in" Then
                a$ = "#0 "
            Else
                a$ = "#0.## "
            End If
            .Text = Format(dblTemp, a$) & cGlobalInfo.DistanceUnits.Display & " from Top"
            If bNotIntialized Then .Rows = .Rows + 1
            intRow = intRow + 1
            .Row = intRow
        Next i%
        
        'Fuel or Additive (if applic) ============================================
        For i% = 1 To cCurTruck.Components.Count
            If cCurTruck.Components(i%).ContentsType = ctFuel Or _
            cCurTruck.Components(i%).ContentsType = ctAdditive Then
                'Tank
                .Col = 0
                .Text = cCurTruck.Components(i%).DisplayName
                'Density
                .Col = 2
                .Text = Format(cCurTruck.Components(i%).Capacity.DensityContents, "0.##")
                'Contents
                .Col = 3
                dblTemp = cCurTruck.Components(i%).Capacity.CurVol * cCurTruck.Components(i%).Capacity.DensityContents
                dblTotalProduct = dblTotalProduct + dblTemp
                dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
                .Text = Format(dblTemp, "#,##0 ") '& cGlobalInfo.MassUnits.Display
                '%Full
                .Col = 4
                dblTemp = cCurTruck.Components(i%).Capacity.CurVol / cCurTruck.Components(i%).Capacity.Volume * 100#
                .Text = Format(dblTemp, "#0") & "%"
                'Fill
                .Col = 5
                If cCurTruck.Components(i%).Capacity.UsesSightGauge Then
                    'Show as sight gauge (percent)
                    .Text = Format(dblTemp, "#0") & "%"
                Else
                    'Show as fill from top
                    dblTemp = cCurTruck.Components(i%).Capacity.CurStkHt * cGlobalInfo.DistanceUnits.Multiplier
                    If cGlobalInfo.DistanceUnits.Display = "in" Then
                        a$ = "#0 "
                    Else
                        a$ = "#0.## "
                    End If
                    .Text = Format(dblTemp, a$) & cGlobalInfo.DistanceUnits.Display & " from Top"
                End If
                If bNotIntialized Then .Rows = .Rows + 1
                intRow = intRow + 1
                .Row = intRow
            End If
        Next i%
        
        'Product Total ==============================================
        .FontItalic = True
        .FontBold = True
        .Col = 1
        .Text = "Total Product"
        'Contents
        .Col = 3
        dblTemp = dblTotalProduct * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .FontItalic = False
        .FontBold = False
        If bNotIntialized Then .Rows = .Rows + 1
        intRow = intRow + 1
        .Row = intRow
        
        'Non-fuel components ==============================================
        For i% = 1 To cCurTruck.Components.Count
            If cCurTruck.Components(i%).ContentsType <> ctFuel And _
            cCurTruck.Components(i%).ContentsType <> ctAdditive And _
            cCurTruck.Components(i%).ContentsType <> ctNone Then
                'Tank
                .Col = 0
                .Text = cCurTruck.Components(i%).DisplayName
                'Density
                .Col = 2
                .Text = Format(cCurTruck.Components(i%).Capacity.DensityContents, "0.##")
                'Contents
                .Col = 3
                dblTemp = cCurTruck.Components(i%).Capacity.CurVol * cCurTruck.Components(i%).Capacity.DensityContents
                dblTotalProduct = dblTotalProduct + dblTemp
                dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
                .Text = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
                '%Full
                .Col = 4
                If cCurTruck.Components(i%).ContentsType = ctOther Then
                    'NOT gassing, water, fuel, or additive
                    If cCurTruck.Components(i%).Capacity.Volume <= 0# Then
                        'No volume capacity listed
                        '.Text = "n/a"
                    Else
                        dblTemp = cCurTruck.Components(i%).Capacity.CurVol / cCurTruck.Components(i%).Capacity.Volume * 100#
                        .Text = Format(dblTemp, "#0") & "%"
                    End If
                Else
                    'This should be a 'normal' volume-type tank
                    dblTemp = cCurTruck.Components(i%).Capacity.CurVol / cCurTruck.Components(i%).Capacity.Volume * 100#
                    .Text = Format(dblTemp, "#0") & "%"
                End If
                'Fill
                .Col = 5
                If cCurTruck.Components(i%).Capacity.DefaultWtContents <> "" Then
                    'Contents indicated by weight which is already shown
                ElseIf cCurTruck.Components(i%).Capacity.DefaultVolContents <> "" Then
                    'Show contents by volume
                    dblTemp = cCurTruck.Components(i%).Capacity.CurVol * cGlobalInfo.VolumeUnits.Multiplier
                    .Text = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
                Else 'show by sightgauge or dist from top
                    If cCurTruck.Components(i%).Capacity.UsesSightGauge Then
                        'Show as sight gauge (percent)
                        .Text = Format(dblTemp, "#0") & "%"
                    Else
                        'Show as fill from top
                        dblTemp = cCurTruck.Components(i%).Capacity.CurStkHt * cGlobalInfo.DistanceUnits.Multiplier
                        If cGlobalInfo.DistanceUnits.Display = "in" Then
                            a$ = "#0 "
                        Else
                            a$ = "#0.## "
                        End If
                        .Text = Format(dblTemp, a$) & cGlobalInfo.DistanceUnits.Display & " from Top"
                    End If
                End If
                If bNotIntialized Then .Rows = .Rows + 1
                intRow = intRow + 1
                .Row = intRow
            End If
        Next i%
        .Redraw = True
    End With
    
    If bNotIntialized Then
        SetGridColumnWidth msgTanks
        With msgTanks
            .Row = 0
            .Col = .Cols - 1
            dblTemp = .CellLeft + .CellWidth + (1 + .Cols) * 70
            msgTanks.Width = dblTemp
        End With
    End If
    Set cAxles = Nothing
    Exit Sub
errHandler:
    ErrorIn "frmMonitor.UpdateTanks"
End Sub


