VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmDefineTruck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Truck"
   ClientHeight    =   7068
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   10560
   Icon            =   "frmDefineTruck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7068
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBodyLocation 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1920
      TabIndex        =   24
      Top             =   1320
      Width           =   852
   End
   Begin VB.CheckBox chkStdMt 
      Alignment       =   1  'Right Justify
      Caption         =   "Standard Mount?"
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
      Left            =   7080
      TabIndex        =   23
      Top             =   1080
      Width           =   1812
   End
   Begin VB.TextBox txtTreadSN 
      Height          =   315
      Left            =   8640
      TabIndex        =   22
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox txtOwner 
      Height          =   315
      Left            =   1920
      TabIndex        =   20
      Top             =   960
      Width           =   3732
   End
   Begin VB.TextBox txtTruckDescription 
      Height          =   315
      Left            =   1920
      TabIndex        =   18
      Top             =   600
      Width           =   3732
   End
   Begin C1SizerLibCtl.C1Tab tabPage 
      Height          =   4812
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   10332
      _cx             =   18224
      _cy             =   8488
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
      Caption         =   "Chassis|Body|Components"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   2
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
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
      Begin C1SizerLibCtl.C1Elastic Page 
         Height          =   4524
         Index           =   2
         Left            =   11064
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   276
         Width           =   10308
         _cx             =   18182
         _cy             =   7980
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
         ForeColorDisabled=   -2147483632
         Caption         =   ""
         Align           =   0
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
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   372
            Left            =   4800
            TabIndex        =   32
            Top             =   3960
            Width           =   1092
         End
         Begin VB.CommandButton cmdAddGeneric 
            Caption         =   "Add Generic"
            Height          =   372
            Left            =   3360
            TabIndex        =   29
            ToolTipText     =   "Add a generic component"
            Top             =   3960
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Frame fraComponent 
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
            Height          =   3972
            Left            =   6240
            TabIndex        =   14
            Top             =   360
            Width           =   3852
            Begin VB.CommandButton cmdModify 
               Caption         =   "Modify"
               Height          =   372
               Left            =   2760
               TabIndex        =   31
               Top             =   3480
               Width           =   972
            End
            Begin VB.TextBox txtIstallationNotes 
               Height          =   852
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   30
               Text            =   "frmDefineTruck.frx":3FBA
               Top             =   2520
               Width           =   3612
            End
            Begin MSFlexGridLib.MSFlexGrid msgComponentInfo 
               Height          =   2292
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   3492
               _ExtentX        =   6160
               _ExtentY        =   4043
               _Version        =   393216
               Rows            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColor       =   -2147483633
               BackColorFixed  =   -2147483626
               BackColorBkg    =   -2147483633
               HighLight       =   0
               GridLines       =   3
               GridLinesFixed  =   3
               AllowUserResizing=   1
               BorderStyle     =   0
               Appearance      =   0
            End
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "ð"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   18
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   2760
            TabIndex        =   13
            Top             =   1920
            Width           =   492
         End
         Begin VB.ListBox lstChosenComponents 
            Height          =   3120
            ItemData        =   "frmDefineTruck.frx":40EA
            Left            =   3360
            List            =   "frmDefineTruck.frx":40F1
            TabIndex        =   12
            Top             =   600
            Width           =   2532
         End
         Begin VB.ListBox lstComponentFiles 
            Height          =   3120
            ItemData        =   "frmDefineTruck.frx":4103
            Left            =   120
            List            =   "frmDefineTruck.frx":410A
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   600
            Width           =   2532
         End
         Begin VB.Label lblComponentFiles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Available Components"
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
            TabIndex        =   9
            Top             =   360
            Width           =   2532
         End
         Begin VB.Label lblChosenComponents 
            BackStyle       =   0  'Transparent
            Caption         =   "Components on Active Truck"
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
            Left            =   3480
            TabIndex        =   11
            Top             =   324
            Width           =   2532
         End
      End
      Begin C1SizerLibCtl.C1Elastic Page 
         Height          =   4524
         Index           =   0
         Left            =   12
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   276
         Width           =   10308
         _cx             =   18182
         _cy             =   7980
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
         ForeColorDisabled=   -2147483647
         Caption         =   ""
         Align           =   0
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
         Begin VB.Frame fraChassis 
            Caption         =   "Chassis Information for Active Truck"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3972
            Left            =   3360
            TabIndex        =   35
            Top             =   360
            Width           =   6732
            Begin VB.TextBox txtWheelBase 
               Height          =   315
               Left            =   5640
               TabIndex        =   39
               Text            =   "444.4"
               ToolTipText     =   "Distance from front-most axle to center of rear tandem (ignores tags)"
               Top             =   3600
               Width           =   612
            End
            Begin VB.CommandButton cmdModifyChassis 
               Caption         =   "Modify"
               Height          =   372
               Left            =   80
               TabIndex        =   36
               Top             =   3540
               Width           =   1452
            End
            Begin MSFlexGridLib.MSFlexGrid msgChassisInfo 
               Height          =   3252
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   6492
               _ExtentX        =   11451
               _ExtentY        =   5736
               _Version        =   393216
               Rows            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColor       =   -2147483633
               BackColorFixed  =   -2147483626
               BackColorBkg    =   -2147483633
               HighLight       =   0
               GridLines       =   3
               GridLinesFixed  =   3
               AllowUserResizing=   1
               BorderStyle     =   0
               Appearance      =   0
            End
            Begin VB.Label lblUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "(in)"
               Height          =   276
               Index           =   1
               Left            =   6300
               TabIndex        =   40
               Top             =   3648
               Width           =   432
            End
            Begin VB.Label lblWheelBase 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Wheelbase"
               Height          =   276
               Left            =   4200
               TabIndex        =   38
               Top             =   3648
               Width           =   1332
            End
         End
         Begin VB.CommandButton cmdUseChassis 
            Caption         =   "ð"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   18
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   2760
            TabIndex        =   33
            Top             =   1680
            Width           =   492
         End
         Begin VB.ListBox lstChassisFiles 
            Height          =   3120
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   2532
         End
         Begin VB.Label lblChooseChassis 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Available Chassis"
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
            TabIndex        =   6
            Top             =   360
            Width           =   2532
         End
      End
      Begin C1SizerLibCtl.C1Elastic Page 
         Height          =   4524
         Index           =   1
         Left            =   10824
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   276
         Width           =   10308
         _cx             =   18182
         _cy             =   7980
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
         Align           =   0
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
         Begin VB.CommandButton cmdUseBody 
            Caption         =   "ð"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   18
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   2760
            TabIndex        =   34
            Top             =   1680
            Width           =   492
         End
         Begin VB.ListBox lstBodyFiles 
            Height          =   3504
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   2532
         End
         Begin MSFlexGridLib.MSFlexGrid msgBodyInfo 
            Height          =   3612
            Left            =   3360
            TabIndex        =   16
            Top             =   600
            Width           =   6732
            _ExtentX        =   11875
            _ExtentY        =   6371
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483633
            BackColorFixed  =   -2147483626
            BackColorBkg    =   -2147483633
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   3
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin VB.Label lblChooseBody 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Available Bodies"
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
            TabIndex        =   7
            Top             =   360
            Width           =   2532
         End
         Begin VB.Label lblgrdBodyInformation 
            BackStyle       =   0  'Transparent
            Caption         =   "Body Information for Active Truck"
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
            Left            =   3480
            TabIndex        =   15
            Top             =   324
            Width           =   6132
         End
      End
   End
   Begin VB.Label lblBodyLocation 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Body Location"
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
      Left            =   240
      TabIndex        =   27
      Top             =   1368
      Width           =   1572
   End
   Begin VB.Label lblBodyLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "(distance from front-most axle)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   1
      Left            =   3240
      TabIndex        =   26
      Top             =   1368
      Width           =   2892
   End
   Begin VB.Label lblUnits 
      BackStyle       =   0  'Transparent
      Caption         =   "(in)"
      Height          =   276
      Index           =   0
      Left            =   2880
      TabIndex        =   25
      Top             =   1368
      Width           =   492
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make necessary changes below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   10452
   End
   Begin VB.Label lblTruckDescription 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Truck Description"
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
      TabIndex        =   17
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label lblOwner 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Owner"
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
      Top             =   960
      Width           =   1692
   End
   Begin VB.Label lblTreadSN 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tread SN"
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
      Left            =   7440
      TabIndex        =   21
      Top             =   600
      Width           =   1092
   End
End
Attribute VB_Name = "frmDefineTruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ALObjects_Body As clsALObjects  'Body objects in lstBodyFiles
Private ALObjects_Chassis As clsALObjects  'Body objects in lstChassisFiles
Private ALObjects_Components As clsALObjects  'Body objects in lstComponentFiles

Private Sub cmdModify_Click()
    Dim ComponentIndex As Integer
    
    ComponentIndex = lstChosenComponents.ListIndex + 1
    frmEditComponent.EditComponent ComponentIndex, Me
    DisplayComponentData
    FillChosenComponentsListBox
    lstChosenComponents.ListIndex = ComponentIndex - 1
End Sub


Private Sub lstChosenComponents_DblClick()
    cmdModify_Click
End Sub


Private Sub Form_Load()
    DisplayTruckData
    DisplayChassisData
    DisplayBodyData
    DisplayComponentData
    FillChosenComponentsListBox
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'Closing window by 'X' is same as pressing 'Done' button
        bHasBeenEdited = True
        bTruckFileDirty = True
        Unload Me
    End If
End Sub

Private Sub DisplayTruckData()
    'Fills-in general truck into
    Dim sLenDispFormat As String
    Dim dblTemp As Double
    
    sLenDispFormat = "#,##0.00 "
    If Int(cGlobalInfo.DistanceUnits.Multiplier * 100) = 3937 Then
        'Inches
        sLenDispFormat = "#,##0.0 "
    End If
    
    txtTruckDescription.Text = cActiveTruck.Description
    txtOwner.Text = cActiveTruck.Owner
    txtTreadSN.Text = cActiveTruck.SN
    chkStdMt.Value = Abs(CLng(cActiveTruck.IsStandardMount))

    'Body Location
    dblTemp = (cActiveTruck.BodyLocation + cActiveTruck.Chassis.TwinSteerSeparation / 2) * cGlobalInfo.DistanceUnits.Multiplier
    txtBodyLocation.Text = Format(dblTemp, sLenDispFormat)
    lblUnits(0).Caption = cGlobalInfo.DistanceUnits.Display
    lblUnits(1).Caption = cGlobalInfo.DistanceUnits.Display
End Sub

Private Sub DisplayChassisData()
    'Fills-in the Chassis tab w/ chassis-related data from the Active Truck
    Dim a$
    Dim b$
    Dim dblTemp As Double
    Dim sLenDispFormat As String
    Dim vRowHt
    Dim i%
        
    sLenDispFormat = "#,##0.000 "
    If Int(cGlobalInfo.DistanceUnits.Multiplier * 100) = 3937 Then
        'Inches
        sLenDispFormat = "#,##0.0 "
    End If
    vRowHt = lblChooseChassis.Height
    
    'Loads active truck data into the various controls on this page
    With msgChassisInfo
        .Clear
        .Rows = 1
        .Enabled = False
        .Visible = False
                
        'Chassis Description
        a$ = "Description"
        b$ = cActiveTruck.Chassis.FullName
        .TextMatrix(0, 0) = a$
        .TextMatrix(0, 1) = b$
        '.AddItem "Description" & vbTab & cActiveTruck.Chassis.FullName
        
        'Chassis Wheelbase
        a$ = "Wheel Base"
        dblTemp = cActiveTruck.Chassis.WB * cGlobalInfo.DistanceUnits.Multiplier
        b$ = Format(dblTemp, sLenDispFormat) & cGlobalInfo.DistanceUnits.Display
        .AddItem a$ & vbTab & b$
        txtWheelBase.Text = Format(dblTemp, sLenDispFormat)
        
        'Tandem spacing
        a$ = "Tandem Spacing"
        dblTemp = cActiveTruck.Chassis.TandemSpacing * cGlobalInfo.DistanceUnits.Multiplier
        b$ = Format(dblTemp, sLenDispFormat) & cGlobalInfo.DistanceUnits.Display
        .AddItem a$ & vbTab & b$
        
        'Empty Wt for Front axle/tire
        dblTemp = cActiveTruck.Chassis.WtFront * cGlobalInfo.MassUnits.Multiplier
        a$ = "Empty Wt Front"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .AddItem a$ & vbTab & b$
        
        'Empty Wt for Rear axle/tire
        dblTemp = cActiveTruck.Chassis.WtRear * cGlobalInfo.MassUnits.Multiplier
        a$ = "Empty Wt Rear"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .AddItem a$ & vbTab & b$
        
        'Manufacturer's Max Load Spec for Front axle/tire
        dblTemp = cActiveTruck.Chassis.WtLimitFront * cGlobalInfo.MassUnits.Multiplier
        a$ = "MfgLimit Front"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .AddItem a$ & vbTab & b$
        
        'Manufacturer's Max Load Spec for Rear axle/tire
        dblTemp = cActiveTruck.Chassis.WtLimitRear * cGlobalInfo.MassUnits.Multiplier
        a$ = "MfgLimit Rear"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .AddItem a$ & vbTab & b$
        
        .ColAlignment(0) = flexAlignLeftCenter 'flexAlignRightCenter
        .ColAlignment(1) = flexAlignLeftCenter
    
        For i% = 0 To .Rows - 1
            .RowHeight(i%) = vRowHt
            .Row = i%
            .Col = 0
            .CellBackColor = vb3DLight
        Next
        SetGridColumnWidth msgChassisInfo
        .Visible = True
        .Enabled = True
    End With
    
    FillListBox cGlobalInfo.ALObjectFolder, lstChassisFiles, xptChassis
End Sub


Private Sub DisplayBodyData()
    'Fills-in the Body tab w/ body-related data from the Active Truck
    Dim a$
    Dim b$
    Dim dblTemp As Double
    Dim dblVol As Double
    Dim sTank As String
    Dim vRowHt
    Dim i%
        
    vRowHt = lblChooseChassis.Height
    
    With msgBodyInfo
        .Clear
        .Rows = 1
        .Enabled = False
        
        'Body Description
        a$ = "Description"
        b$ = cActiveTruck.Body.FullName
        .TextMatrix(0, 0) = a$
        .TextMatrix(0, 1) = b$
        
        dblVol = 0
        For i% = 1 To cActiveTruck.Body.Tanks.Count
        
            'Tank Name
            sTank = cActiveTruck.Body.Tanks(i%).DisplayName & ": "
        
            'Tank Type
            a$ = sTank & "Type"
            b$ = cActiveTruck.Body.Tanks(i%).TankTypeString
            .AddItem a$ & vbTab & b$
        
            'Tank Volume
            dblTemp = cActiveTruck.Body.Tanks(i%).Volume * cGlobalInfo.VolumeUnits.Multiplier
            a$ = sTank & "Capacity"
            b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
            .AddItem a$ & vbTab & b$
        
        
            dblVol = dblVol + cActiveTruck.Body.Tanks(i%).Volume
        Next
        
        'Tank Volume
        dblTemp = dblVol * cGlobalInfo.VolumeUnits.Multiplier
        a$ = "TOTAL Capacity"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
        .AddItem a$ & vbTab & b$
    
        .ColAlignment(0) = flexAlignLeftCenter 'flexAlignRightCenter
        .ColAlignment(1) = flexAlignLeftCenter
        
        For i% = 0 To .Rows - 1
            .RowHeight(i%) = vRowHt
            .Row = i%
            .Col = 0
            .CellBackColor = vb3DLight
        Next
        SetGridColumnWidth msgBodyInfo
        .Enabled = True
    End With
    
    'Only advanced users can change the body type.  Others should use templates.
    If bCanCreateALObjects Then
        lstBodyFiles.Visible = True
        FillListBox cGlobalInfo.ALObjectFolder, lstBodyFiles, xptBody
        cmdUseBody.Visible = True
        lblChooseBody.Visible = True
        msgBodyInfo.Left = 3360
        msgBodyInfo.Width = 6732
        lblgrdBodyInformation.Left = 3480
    Else
        lstBodyFiles.Visible = False
        cmdUseBody.Visible = False
        lblChooseBody.Visible = False
        msgBodyInfo.Left = 120
        msgBodyInfo.Width = 9972
        lblgrdBodyInformation.Left = 240
    End If

End Sub

Private Sub DisplayComponentData()
    Dim a$
    Dim b$
    Dim dblTemp As Double
    Dim dblVol As Double
    Dim sTank As String
    Dim vRowHt

    FillListBox cGlobalInfo.ALObjectFolder, lstComponentFiles, xptComponent
    lstChosenComponents.ListIndex = 0
End Sub

Private Sub cmdUseBody_Click()
    'Replace Active body with selected body
    Dim lIndex As Long
    Dim ALObject As clsALObject
    Dim sFile As String
    
    lIndex = lstBodyFiles.ListIndex
    If lIndex < 0 Then Exit Sub
    'Get the index that links the ALObject
    lIndex = lstBodyFiles.ItemData(lstBodyFiles.ListIndex)
    
    Set ALObject = ALObjects_Body(CStr(lIndex))
    sFile = ALObject.File
    
    If ReplaceBody(cActiveTruck, sFile) Then
        'Succsessful. Refresh display.
        DisplayBodyData
    End If
End Sub


Private Sub lstChosenComponents_Click()
    'Show information related to chosen component
    Dim lIndex As Integer
    Dim cComponent As New clsComponent
    Dim i%
    Dim a$
    Dim b$
    Dim dblTemp As Double
    Dim sLenDispFormat As String
    Dim vRowHt
    Dim lWidth As Long
    Dim msgMaxHt As Long
    
    msgMaxHt = 2292 'grid can be no taller than this value
    vRowHt = lblChooseChassis.Height
    
    sLenDispFormat = "#,##0.00 "
    If Int(cGlobalInfo.DistanceUnits.Multiplier * 100) = 3937 Then
        'Inches
        sLenDispFormat = "#,##0 "
    End If
    
    lIndex = lstChosenComponents.ListIndex
    If lIndex < 0 Then GoTo CleanUp
    
    Set cComponent = cActiveTruck.Components(lIndex + 1) 'Component collection is not zero-based
    
    With msgComponentInfo
        .Enabled = False
        .Visible = False
        .Clear
        .Height = msgMaxHt
        .Rows = 1
        
        'Component Description
        a$ = "Description"
        b$ = cComponent.FullName
        .TextMatrix(0, 0) = a$
        .TextMatrix(0, 1) = b$
        
        'Component Empty Wt
        dblTemp = cComponent.EmptyWeight * cGlobalInfo.MassUnits.Multiplier
        a$ = "Empty Wt"
        b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .AddItem a$ & vbTab & b$
        
        'Location Ref
        a$ = "Loc. Reference"
        b$ = cComponent.LocationReferenceString
        .AddItem a$ & vbTab & b$
        
        'Offset
        a$ = "Offset fm Reference"
        dblTemp = cComponent.Offset * cGlobalInfo.DistanceUnits.Multiplier
        b$ = Format(dblTemp, sLenDispFormat) & cGlobalInfo.DistanceUnits.Display
        .AddItem a$ & vbTab & b$
        
        'Placements
        a$ = "Placement"
        b$ = cComponent.PlacementString
        .AddItem a$ & vbTab & b$
        
        'Contents Type
        a$ = "Contents Type"
        b$ = cComponent.ContentsTypeString
        .AddItem a$ & vbTab & b$
        
        'Show contents
        If cComponent.ContentsType = ctOther Then
            If cComponent.Capacity.DefaultVolContents <> "" Then
                'Contents are by volume
                dblTemp = cComponent.Capacity.DefaultVolContents * cGlobalInfo.VolumeUnits.Multiplier
                a$ = "Contents Volume"
                b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
                .AddItem a$ & vbTab & b$
            ElseIf cComponent.Capacity.DefaultWtContents <> "" Then
                'Contents are by Mass
                dblTemp = cComponent.Capacity.DefaultWtContents * cGlobalInfo.MassUnits.Multiplier
                a$ = "Contents Wt"
                b$ = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
                .AddItem a$ & vbTab & b$
            End If
        End If
        
        .ColAlignment(0) = flexAlignLeftCenter 'flexAlignRightCenter
        .ColAlignment(1) = flexAlignLeftCenter
    
        For i% = 0 To .Rows - 1
            .RowHeight(i%) = vRowHt
            .Row = i%
            .Col = 0
            .CellBackColor = vb3DLight
        Next
        
        SetGridColumnWidth msgComponentInfo
        
        .Enabled = True
        .Visible = True
    End With
    
    fraComponent.Caption = "Component (" & cComponent.DisplayName & ")"
    txtIstallationNotes.Text = cComponent.InstallationNotes
    
CleanUp:
    Set cComponent = Nothing
End Sub


Private Sub lstChosenComponents_KeyUp(KeyCode As Integer, Shift As Integer)
    'Remove the component highlighted in the lstChosenComponents listbox
    'Show information related to chosen component
    Dim lIndex As Integer
    
    If KeyCode = vbKeyDelete Then
        lIndex = lstChosenComponents.ListIndex
        If lIndex < 0 Then Exit Sub 'Nothing highlighted for deletion
        
        cActiveTruck.Components.Remove (lIndex + 1)
        FillChosenComponentsListBox
        lstChosenComponents.SetFocus
        If lIndex <= lstChosenComponents.ListCount - 1 Then
            lstChosenComponents.ListIndex = lIndex
        Else
            lstChosenComponents.ListIndex = lIndex - 1
        End If
    End If
End Sub

Private Sub msgComponentInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Set TooltipText as value
    msgComponentInfo.ToolTipText = msgComponentInfo.TextMatrix(msgComponentInfo.MouseRow, 1)
End Sub

Private Sub txtBodyLocation_Validate(Cancel As Boolean)
    'Change body location
    Dim dblTemp As Double
    
    With txtBodyLocation
        If Trim$(.Text) <> "" And IsNumber(.Text) Then
            dblTemp = (CDbl(.Text) / cGlobalInfo.DistanceUnits.Multiplier) - cActiveTruck.Chassis.TwinSteerSeparation / 2
            cActiveTruck.BodyLocation = dblTemp
        Else
            MsgBox "You must enter a valid number for Body Location"
            Cancel = True
        End If
    End With
End Sub

Private Sub txtOwner_Validate(Cancel As Boolean)
    With txtOwner
        If Trim$(.Text) <> "" Then
            cActiveTruck.Owner = Trim$(.Text)
        Else
            MsgBox "You must enter a value for Owner"
            Cancel = True
        End If
    End With
End Sub

Private Sub txtTreadSN_Validate(Cancel As Boolean)
    With txtTreadSN
        If Trim$(.Text) <> "" Then
            cActiveTruck.SN = Trim$(.Text)
        Else
            MsgBox "You must enter a value for Tread SN"
            Cancel = True
        End If
    End With
End Sub

Private Sub txtTruckDescription_Validate(Cancel As Boolean)
    With txtTruckDescription
        If Trim$(.Text) <> "" Then
            cActiveTruck.Description = Trim$(.Text)
        Else
            MsgBox "You must enter a value for Truck Description"
            Cancel = True
        End If
    End With
End Sub

Private Sub chkStdMt_Click()
    If chkStdMt.Value = vbChecked Then
        cActiveTruck.IsStandardMount = True
    Else
        cActiveTruck.IsStandardMount = False
    End If
End Sub

Private Sub FillListBox(sPath As String, ctlListBox As ListBox, ObjectType As ALObjectType)
    'Loads up the list box with all relevant objects in given path
    Dim cALObjects As New clsALObjects
    Dim cALObject As clsALObject
    Dim i%
    
    ctlListBox.Clear
    Set cALObjects = ALObjectsInPath(sPath, ObjectType)
    i% = 0
    If TypeName(cALObjects) <> "Nothing" Then
        For Each cALObject In cALObjects
            ctlListBox.AddItem cALObject.DisplayName, i%
            ctlListBox.ItemData(i%) = cALObject.Index
            i% = i% + 1
        Next
    End If
    
    Select Case ObjectType
    Case xptBody
        Set ALObjects_Body = cALObjects
    Case xptChassis
        Set ALObjects_Chassis = cALObjects
    Case xptComponent
        Set ALObjects_Components = cALObjects
    End Select
End Sub

Private Sub FillChosenComponentsListBox()
    'This routine should be called at form load time,
    ' and any time that the list of components changes
    Dim i%
    Dim cComponent As New clsComponent

    With lstChosenComponents
        .Enabled = False
        .Visible = False
        .Clear
        
        For i% = 1 To cActiveTruck.Components.Count
            Set cComponent = cActiveTruck.Components(i%)
            .AddItem cComponent.DisplayName, i% - 1
        Next
        
        .Visible = True
        .Enabled = True
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    Set cComponent = Nothing
End Sub

Private Sub lstChassisFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LstBoxToolTip lstChassisFiles, x, y
End Sub

Private Sub lstBodyFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LstBoxToolTip lstBodyFiles, x, y
End Sub

Private Sub lstComponentFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LstBoxToolTip lstComponentFiles, x, y
End Sub


Private Sub LstBoxToolTip(ctlListBox As ListBox, x As Single, y As Single)
    'Sets the ToolTipText equal to the ALObject.FullName
    Dim oldFont As StdFont, itemIndex As Long
    
    With ctlListBox.Parent
        Set oldFont = .Font
        Set .Font = ctlListBox.Font
        ' determine which element the mouse is on
        itemIndex = y \ .TextHeight("A") + ctlListBox.TopIndex
        ' restore fonts
        Set .Font = oldFont
    End With
    
    ' set the tooltip to the current item's string
    If itemIndex < ctlListBox.ListCount Then
        itemIndex = ctlListBox.ItemData(itemIndex)
        Select Case ctlListBox.Name
        Case "lstChassisFiles"
            ctlListBox.ToolTipText = ALObjects_Chassis(CStr(itemIndex)).FullName
        Case "lstBodyFiles"
            ctlListBox.ToolTipText = ALObjects_Body(CStr(itemIndex)).FullName
        Case "lstComponentFiles"
            ctlListBox.ToolTipText = ALObjects_Components(CStr(itemIndex)).FullName
        Case Else
            ctlListBox.ToolTipText = ""
        End Select
    Else
        ctlListBox.ToolTipText = ""
    End If
End Sub

Private Sub cmdUseChassis_Click()
    'Replace chassis with selected chassis
    Dim lIndex As Long
    Dim ALObject As clsALObject
    Dim sFile As String
    
    lIndex = lstChassisFiles.ListIndex
    If lIndex < 0 Then Exit Sub
    'Get the index that links the ALObject
    lIndex = lstChassisFiles.ItemData(lstChassisFiles.ListIndex)
    
    Set ALObject = ALObjects_Chassis(CStr(lIndex))
    sFile = ALObject.File
    
    If ReplaceChassis(cActiveTruck, sFile) Then
        'Succsessful. Refresh display.
        DisplayChassisData
        FillChosenComponentsListBox
    End If
End Sub

Private Sub cmdRemove_Click()
    'Remove the component highlighted in the lstChosenComponents listbox
    'Show information related to chosen component
    Dim lIndex As Integer
    
    lIndex = lstChosenComponents.ListIndex
    If lIndex < 0 Then Exit Sub 'Nothing highlighted for deletion
    
    cActiveTruck.Components.Remove (lIndex + 1)
    FillChosenComponentsListBox
    lstChosenComponents.SetFocus
    If lIndex <= lstChosenComponents.ListCount - 1 Then
        lstChosenComponents.ListIndex = lIndex
    Else
        lstChosenComponents.ListIndex = lIndex - 1
    End If
End Sub

Private Sub cmdModifyChassis_Click()
    frmEditChassis.Show vbModal, Me
    DisplayChassisData 'update screen w/ any changes
End Sub


Private Sub cmdAdd_Click()
    'Add selected component to Active truck
    Dim lIndex As Long
    Dim ALObject As clsALObject
    Dim sFile As String
    
    lIndex = lstComponentFiles.ListIndex
    If lIndex < 0 Then Exit Sub
    'Get the index that links the ALObject
    lIndex = lstComponentFiles.ItemData(lstComponentFiles.ListIndex)
    
    Set ALObject = ALObjects_Components(CStr(lIndex))
    sFile = ALObject.File
    
    If LoadComponent(cActiveTruck, sFile) Then
        'Succsessful. Refresh display.
        DisplayChassisData
        FillChosenComponentsListBox
    End If
End Sub


Private Sub txtWheelBase_Validate(Cancel As Boolean)
    If Not IsNumber(txtWheelBase.Text) Then
        Beep
        MsgBox "Please enter a valid number for Wheelbase", vbExclamation
        Cancel = True
    Else
        'Update wheelbase & displays of wheelbase
        cActiveTruck.Chassis.WB = CDbl(txtWheelBase.Text) / cGlobalInfo.DistanceUnits.Multiplier
        DisplayChassisData
    End If

End Sub
