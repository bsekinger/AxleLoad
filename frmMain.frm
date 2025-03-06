VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unibody AxleLoad"
   ClientHeight    =   7344
   ClientLeft      =   120
   ClientTop       =   804
   ClientWidth     =   8664
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7344
   ScaleWidth      =   8664
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCalculate 
      BackColor       =   &H00A4A3A3&
      Caption         =   "Find best loading for current tank configuration"
      Height          =   372
      Index           =   1
      Left            =   5640
      TabIndex        =   39
      Top             =   6120
      Width           =   2652
   End
   Begin VB.OptionButton optCalculate 
      BackColor       =   &H00A4A3A3&
      Caption         =   "Find maximum loading capacity (may require re-configuring tanks)"
      Height          =   372
      Index           =   0
      Left            =   5640
      TabIndex        =   38
      Top             =   5600
      Value           =   -1  'True
      Width           =   2652
   End
   Begin VB.Frame fraSolve 
      BackColor       =   &H00A4A3A3&
      Caption         =   "Find Solution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      Left            =   4320
      TabIndex        =   30
      Top             =   5160
      Width           =   4092
      Begin VB.CommandButton cmdAutoTankFill 
         Caption         =   "Automatically"
         Height          =   300
         Left            =   2640
         TabIndex        =   42
         Top             =   1560
         Width           =   1332
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Solve"
         Height          =   852
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lblAutoTankFill 
         BackStyle       =   0  'Transparent
         Caption         =   "Set fuel, water, gas tank levels:"
         Height          =   276
         Left            =   240
         TabIndex        =   43
         Top             =   1580
         Width           =   2292
      End
      Begin VB.Line lin 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   3960
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.Frame fraCriteria 
      BackColor       =   &H00A4A3A3&
      Caption         =   "Loading Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2532
      Left            =   4320
      TabIndex        =   26
      Top             =   2520
      Width           =   4092
      Begin VB.TextBox txtProductLimit 
         BackColor       =   &H00A4A3A3&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   36
         Text            =   "15000"
         Top             =   1970
         Width           =   732
      End
      Begin VB.CheckBox chkLimitProduct 
         BackColor       =   &H00A4A3A3&
         Caption         =   "Limit Product to:"
         Height          =   252
         Left            =   240
         TabIndex        =   35
         Top             =   2000
         Width           =   1692
      End
      Begin VB.TextBox txtRatio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   612
      End
      Begin VB.CheckBox chkOnRoad 
         BackColor       =   &H00A4A3A3&
         Caption         =   "Find loading that is road-legal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   3732
      End
      Begin VB.Label lblProductLimit 
         BackStyle       =   0  'Transparent
         Caption         =   "lbs"
         Height          =   252
         Left            =   2520
         TabIndex        =   37
         Top             =   2040
         Width           =   492
      End
      Begin VB.Label lblMaxEmul 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Emulsion = "
         Height          =   276
         Left            =   240
         TabIndex        =   34
         Top             =   1680
         Width           =   3612
      End
      Begin VB.Label lblMaxAN 
         BackStyle       =   0  'Transparent
         Caption         =   "Max AN ="
         Height          =   252
         Left            =   240
         TabIndex        =   33
         Top             =   1440
         Width           =   3732
      End
      Begin VB.Label lblRatio 
         BackStyle       =   0  'Transparent
         Caption         =   "% Emulsion Desired"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   936
         TabIndex        =   29
         Top             =   504
         Width           =   2772
      End
   End
   Begin VB.Frame fraTruckInfo 
      Caption         =   "Truck Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   4092
      Begin VB.Label lblVersionLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Software Version:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   0
         TabIndex        =   41
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   1320
         TabIndex        =   40
         Top             =   1200
         Width           =   2652
      End
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         Height          =   252
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   2652
      End
      Begin VB.Label lblFileNameLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Source File:"
         Height          =   252
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   972
      End
      Begin VB.Label lblSN 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         Height          =   276
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   2652
      End
      Begin VB.Label lblBody 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         Height          =   276
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   2652
      End
      Begin VB.Label lblChassis 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         Height          =   276
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   2652
      End
      Begin VB.Label lblChassisLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Chassis:"
         Height          =   252
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   972
      End
      Begin VB.Label lblSNLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tread SN:"
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   972
      End
      Begin VB.Label lblBodyLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Body Style:"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   972
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Component Contents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   3972
      Begin C1SizerLibCtl.C1Tab tabComponent 
         Height          =   1572
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   3492
         _cx             =   6159
         _cy             =   2773
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
         Caption         =   "Comp 1"
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
         Begin VB.Frame fraHolder 
            BorderStyle     =   0  'None
            Height          =   1284
            Left            =   12
            TabIndex        =   55
            Top             =   276
            Width           =   3468
            Begin VB.TextBox txtDefault 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   57
               Text            =   "xxx"
               Top             =   360
               Width           =   972
            End
            Begin VB.TextBox txtCurVal 
               Height          =   315
               Left            =   1200
               TabIndex        =   56
               Top             =   360
               Width           =   1092
            End
            Begin VB.Label lblCompUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "Units"
               Height          =   276
               Index           =   1
               Left            =   2400
               TabIndex        =   63
               Top             =   840
               Width           =   732
            End
            Begin VB.Label lblVolume 
               BackStyle       =   0  'Transparent
               Caption         =   "XXXX"
               Height          =   276
               Left            =   1200
               TabIndex        =   62
               Top             =   840
               Width           =   972
            End
            Begin VB.Label lblCapacity 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Capacity ="
               Height          =   276
               Left            =   120
               TabIndex        =   61
               Top             =   840
               Width           =   972
            End
            Begin VB.Label lblDefault 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Default"
               Height          =   276
               Left            =   120
               TabIndex        =   60
               Top             =   120
               Width           =   972
            End
            Begin VB.Label lblCurVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Current Value"
               Height          =   276
               Left            =   1200
               TabIndex        =   59
               Top             =   120
               Width           =   1092
            End
            Begin VB.Label lblCompUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "Units"
               Height          =   276
               Index           =   0
               Left            =   2400
               TabIndex        =   58
               Top             =   480
               Width           =   732
            End
         End
      End
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
      Height          =   1812
      Left            =   4320
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   4212
      Begin C1SizerLibCtl.C1Tab tabTagDisplay 
         Height          =   1452
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   3732
         _cx             =   6583
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
            ScaleWidth      =   3708
            TabIndex        =   45
            Top             =   276
            Width           =   3708
            Begin VB.TextBox txtTag 
               Height          =   315
               Left            =   120
               TabIndex        =   46
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
               TabIndex        =   53
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
               TabIndex        =   52
               Top             =   120
               Width           =   1332
            End
            Begin VB.Label lblMassUnits 
               BackStyle       =   0  'Transparent
               Caption         =   "kg"
               Height          =   276
               Index           =   0
               Left            =   1200
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
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
               TabIndex        =   48
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
               TabIndex        =   47
               ToolTipText     =   "Provides estimate for tag air pressure to achieve indicated ""Applied Force"""
               Top             =   408
               Width           =   972
            End
         End
      End
   End
   Begin VB.Frame fraTankConfiguration 
      Caption         =   "Current Tank Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   3972
      Begin VB.ComboBox cboTankConfig 
         Height          =   288
         Index           =   4
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1800
         Width           =   1212
      End
      Begin VB.ComboBox cboTankConfig 
         Height          =   288
         Index           =   3
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1320
         Width           =   1212
      End
      Begin VB.ComboBox cboTankConfig 
         Height          =   288
         Index           =   2
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1212
      End
      Begin VB.ComboBox cboTankConfig 
         Height          =   288
         Index           =   1
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label lblDensity 
         Caption         =   "1.2"
         Height          =   276
         Index           =   4
         Left            =   2880
         TabIndex        =   25
         Top             =   1836
         Width           =   852
      End
      Begin VB.Label lblDensity 
         Caption         =   "1.2"
         Height          =   276
         Index           =   3
         Left            =   2880
         TabIndex        =   24
         Top             =   1356
         Width           =   852
      End
      Begin VB.Label lblDensity 
         Caption         =   "1.2"
         Height          =   276
         Index           =   2
         Left            =   2880
         TabIndex        =   23
         Top             =   876
         Width           =   852
      End
      Begin VB.Label lblDensity 
         Caption         =   "1.2"
         Height          =   276
         Index           =   1
         Left            =   2880
         TabIndex        =   22
         Top             =   396
         Width           =   852
      End
      Begin VB.Label lblDensity 
         Caption         =   "Density"
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
         Left            =   2880
         TabIndex        =   21
         Top             =   120
         Width           =   732
      End
      Begin VB.Label lblTankConfig 
         BackStyle       =   0  'Transparent
         Caption         =   "Tank D"
         Height          =   276
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label lblTankConfig 
         BackStyle       =   0  'Transparent
         Caption         =   "Tank C"
         Height          =   276
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1332
      End
      Begin VB.Label lblTankConfig 
         BackStyle       =   0  'Transparent
         Caption         =   "Tank B"
         Height          =   276
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label lblTankConfig 
         BackStyle       =   0  'Transparent
         Caption         =   "Tank A"
         Height          =   276
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1332
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3720
      Top             =   1920
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H00A4A3A3&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4692
      Left            =   4200
      TabIndex        =   32
      Top             =   2520
      Width           =   4332
   End
   Begin VB.Label lblTruck 
      Caption         =   "[No File Loaded]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8412
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mSaveAsDelimiter 
         Caption         =   "-"
      End
      Begin VB.Menu mSetTruckFolder 
         Caption         =   "Set Truck Folder"
      End
      Begin VB.Menu mExitDelimiter 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mConfiguration 
      Caption         =   "&Configuration"
      Begin VB.Menu mEditTruck 
         Caption         =   "&Edit Truck"
      End
      Begin VB.Menu mDrawTruck 
         Caption         =   "&Draw Truck"
      End
      Begin VB.Menu mShopReport 
         Caption         =   "&Shop Report"
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "&Tools"
      Begin VB.Menu mDensity 
         Caption         =   "Change Densities"
      End
      Begin VB.Menu mNoLowFill 
         Caption         =   "Avoid Tank Fill < 5%"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu mManualFill 
         Caption         =   "Manually Fill Tanks"
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mAdvanced 
         Caption         =   "Advanced"
         Begin VB.Menu mRelationships 
            Caption         =   "Change Fill Relationships"
         End
         Begin VB.Menu mAdjustTag 
            Caption         =   "Calibrate Tag Pressure"
         End
         Begin VB.Menu mAdjustEmpty 
            Caption         =   "Adjust Empty Weight"
         End
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mTruckInfo 
         Caption         =   "Truck Info"
      End
      Begin VB.Menu m5 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About Unibody AxleLoad"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bAlreadyStarted As Boolean
Private strCurFile As String
Private bFindBest As Boolean
Private bConfigurable As Boolean 'TRUE=Current truck has dual-use tanks
Private dblProductLimit As Double
Private CommandLineFile As String 'name of truck file passed via command line

Private Const BLUE_BKGD As Long = &HA4A3A3      '&HBF6640


Private Sub chkLimitProduct_Click()
    If chkLimitProduct.Value = 1 Then
        txtProductLimit.Enabled = True
        txtProductLimit.BackColor = vbWindowBackground
    Else
        txtProductLimit.Enabled = False
        txtProductLimit.BackColor = BLUE_BKGD
    End If
End Sub

Private Sub cmdAutoTankFill_Click()
    frmAutoTankFilling.Options cActiveTruck
End Sub

Private Sub Form_Activate()
    Dim sFile As String
    Dim sErr As String
    Dim a
    On Error GoTo errHandler
    
    If bInHouseVersion Then
        mSaveAs.Visible = True
        mSaveAsDelimiter.Visible = True
        mConfiguration.Visible = True
    Else
        mSaveAs.Visible = False
        mSaveAsDelimiter.Visible = False
        mConfiguration.Visible = False
    End If
    
    If Not bAlreadyStarted Then
        'Only do this at startup
        bTruckFileDirty = False
        
        Set cActiveTruck = New clsTruck
        'Read GlobalInfo
        sFile = AddBackslash(App.Path) & "GlobalInfo.xml"
        sErr = ReadGlobalInfo(sFile)
        If sErr <> "" Then
            MsgBox "Error loading Truck's XML File" & vbCrLf & sErr
            Exit Sub
        End If
        
        'Now check for command-line arguments
        If Len(Command) <= 4 Then
            CommandLineFile = "" 'no file on command line
        Else
            CommandLineFile = Trim$(Replace(Command, Chr$(34), "")) 'eliminate quatation marks
            If UCase$(Right$(CommandLineFile, 4)) <> ".XML" Then
                CommandLineFile = "" 'no file on command line
            End If
        End If
        
        'Loop until user select a file
        Do
            mOpen_Click 'Make user get a file
            If cActiveTruck.SN = "" Then
                a = MsgBox("You must select a truck file.", vbRetryCancel)
                If a = vbCancel Then
                    Unload Me
                    Exit Sub
                End If
            End If
        Loop Until cActiveTruck.SN <> ""
        bAlreadyStarted = True
    End If
    InitForm
    mSave.Enabled = bTruckFileDirty 'Save only accessible if changes were made
    Exit Sub
errHandler:
    ErrorIn "frmMain.Form_Activate"
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim F As Form
    On Error GoTo errHandler
    
    For Each F In Forms
        'Unload all forms except "Me" which will be unloaded @ sub exit
        If Not (F Is Me) Then
            Unload F
        End If
    Next
    Exit Sub
errHandler:
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel
End Sub

Private Sub mAbout_Click()
    frmAbout.Show
End Sub

Private Sub mAdjustEmpty_Click()
    frmAdjustWeight.Edit cActiveTruck
End Sub

Private Sub mDensity_Click()
    frmDensity.Edit cActiveTruck
End Sub


'------------------------------------------------------------------------------------
#If TREADVERSION = 1 Then ' Only compiled in Tread Version of the program -----------------
Private Sub mEditTruck_Click()
    frmDefineTruck.Show 'vbModal, Me
    'Setup the screen
    InitForm
End Sub
Private Sub mDrawTruck_Click()
    frmReport.ShowDrawing cActiveTruck
    'frmReport.PrintDXF cActiveTruck, "D:\My Documents\VB_Progs\Axleload_Unibody\Tread Version\~TruckDrawing.bmp"
End Sub
Private Sub mShopReport_Click()
    frmReport.ShopReport cActiveTruck
End Sub
#End If
'------------------------------------------------------------------------------------

Private Sub mManualFill_Click()
    frmMonitor.ManualFill cActiveTruck
End Sub

Private Sub mNoLowFill_Click()
    mNoLowFill.Checked = Not mNoLowFill.Checked
    cGlobalInfo.AvoidLowFill = mNoLowFill.Checked
    bTruckFileDirty = True
End Sub

Private Sub mRelationships_Click()
    frmFillRelationships.Edit cActiveTruck
End Sub

Private Sub mAdjustTag_Click()
    frmTagPressure.Edit cActiveTruck
End Sub

Private Sub mSetTruckFolder_Click()
    Dim sMsg As String
    Dim sFile As String
    On Error GoTo errHandler
    
    sMsg = "Current Truck Folder = " & Chr$(34) & cGlobalInfo.TruckFolder & Chr$(34) & _
           vbCrLf & vbCrLf & "Click OK to change."

    If MsgBox(sMsg, vbOKCancel, "Change Default Truck Folder") = vbOK Then
        sFile = BrowseFolders(Me.hwnd, "Select New Truck Folder")
        If sFile = "" Then
            'user cancelled, so exit
            Exit Sub
        End If
        cGlobalInfo.TruckFolder = sFile
        'Save GlobalInfo
        sFile = AddBackslash(App.Path) & "GlobalInfo.xml"
        SaveGlobalInfo sFile
    End If
    Exit Sub
errHandler:
    ErrorIn "frmMain.mSetTruckFolder_Click"
End Sub

Private Sub mTruckInfo_Click()
    frmTruckInfo.Edit strCurFile, cActiveTruck
End Sub

Private Sub mExit_Click()
    Unload Me
End Sub


Private Sub mOpen_Click()
    Dim sErr As String
    Dim msg As String
    On Error GoTo errHandler
    
    'Alert user if current file is unsaved (and save-able)
    If TypeName(cActiveTruck.CreateVersion) <> "Nothing" Then
        If bTruckFileDirty And Not (cActiveTruck.CreateVersion.Major > App.Major) _
           And Not (cActiveTruck.CreateVersion.Major = App.Major _
                 And cActiveTruck.CreateVersion.Minor > App.Minor) Then
            
            msg = "You have not saved the file you were working on." & vbCrLf & _
                  "Press 'Cancel' if you would like a chance to save the current file."
            If MsgBox(msg, vbOKCancel Or vbQuestion, "Continue without saving?") = vbCancel Then
                ' User selected Cancel
                Exit Sub
            End If
        End If
    End If
    
    With dlgFile
        .CancelError = True
        .InitDir = cGlobalInfo.TruckFolder
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        .Filter = "Truck Files|Truck*.xml"
        On Error Resume Next
        If CommandLineFile = "" Then
            .ShowOpen 'Show the OpenFile dialog
        Else
            .FileName = CommandLineFile
            CommandLineFile = ""
        End If
        'If User selected a file, open it
        If Err = 0 Then
            'Read Truck
            On Error GoTo errHandler
            strCurFile = .FileName
            sErr = LoadTruck(strCurFile, cActiveTruck)
            If sErr <> "" Then
                MsgBox "Error loading File: '" & strCurFile & "'" & vbCrLf & sErr
                strCurFile = ""
                Exit Sub
            End If
            
            'See if this truck was created on newer version of software
            If cActiveTruck.CreateVersion.Major > App.Major Then
                'Loaded a file made on a MUCH NEWER version of software
                msg = "This file was created on a much newer version of this software. " & vbCrLf & _
                        "Please contact Tread about upgrading to the latest version."
                MsgBox msg, vbExclamation, "Cannot Load file"
                Exit Sub
            ElseIf cActiveTruck.CreateVersion.Major < App.Major Then
                'Loaded a file made on a MUCH OLDER version of software
                msg = "This file was created on a much older version of this software. " & vbCrLf & _
                        "Although the file will be loaded and should work, it is suggested " & vbCrLf & _
                        "that, for this particular file, you use an older version (" & _
                        cActiveTruck.CreateVersion.Major & ".x" & ")" & vbCrLf & _
                        "of the AxleLoad program."
                MsgBox msg, vbInformation, "Notice"
            ElseIf cActiveTruck.CreateVersion.Major = App.Major _
             And cActiveTruck.CreateVersion.Minor > App.Minor Then
                'Loaded a file made on a slightly NEWER version of software
                msg = "This file was created on a newer version of this software. " & vbCrLf & _
                        "Although the file WILL be loaded and should work, it is" & vbCrLf & _
                        "suggested that you upgrade to the latest version of the" & vbCrLf & _
                        "AxleLoad software in order to take advantage of new functionality." & vbCrLf & _
                        "(Contact Tread for information)"
                MsgBox msg, vbInformation, "Notice"
            End If
            
            
            sErr = SetComponentDensities(cActiveTruck)
            If sErr <> "" Then
                MsgBox "Error setting Component Densities" & vbCrLf & sErr
                Exit Sub
            End If
            sErr = SetProductDensities(cActiveTruck)
            If sErr <> "" Then
                MsgBox "Error setting Product Densities" & vbCrLf & sErr
                Exit Sub
            End If
            InitializeTanks cActiveTruck
            'Setup the screen
            InitForm
            
            bTruckFileDirty = False
        End If
    End With
    Exit Sub
errHandler:
    ErrorIn "frmMain.mOpen_Click"
End Sub


Private Sub mSave_Click()
    Dim stMsg As String
    On Error GoTo errHandler
    
    'Prevent saving of files created on newer software
    If cActiveTruck.CreateVersion.Major = App.Major And _
           cActiveTruck.CreateVersion.Minor > App.Minor Then
        stMsg = "This file was created on a newer version of this software. " & vbCrLf & _
                "Please contact Tread about upgrading to the latest version."
        MsgBox stMsg, vbExclamation, "Cannot save file"
        Exit Sub
    End If
    
    'Warn user of what they are about to do
    If bHasBeenEdited Then
        stMsg = "This file has been edited." & vbCrLf & _
                "Do you really want to save it with the same name?" & vbCrLf & vbCrLf & _
                "Press Cancel and choose 'Save As' to give this file another name."
        If MsgBox(stMsg, vbOKCancel, "Save Edited File?") = vbCancel Then
            'User chose not to continue
            Exit Sub
        End If
     End If
    
    Call SaveTruck(strCurFile, cActiveTruck)
    
    bHasBeenEdited = False
    bTruckFileDirty = False
    mSave.Enabled = bTruckFileDirty 'Save only accessible if changes were made

    Exit Sub
errHandler:
    ErrorIn "frmMain.mSave_Click"
End Sub

'------------------------------------------------------------------------------------
#If TREADVERSION = 1 Then ' Only compiled in Tread Version of the program -----------------
Private Sub mSaveAs_Click()
    'Save the modified truck to a new file
    Dim sMsg As String
    
    sMsg = SaveTruckAs(cActiveTruck, strCurFile)
    If sMsg <> "" Then
        MsgBox sMsg, vbExclamation, "Failed to Save Truck File"
        Exit Sub
    End If
    
    InitForm
    mSave.Enabled = bTruckFileDirty 'Save only accessible if changes were made
End Sub
#End If
'------------------------------------------------------------------------------------


Private Sub txtProductLimit_Change()
    txtProductLimit_Validate False
End Sub

Private Sub txtProductLimit_Validate(Cancel As Boolean)
    'Check Tag value.  If OK, convert to internal units and save value
    Dim dblNum As Double
    Dim dblMax As Double 'Max value (in approp. eng. units)
    Dim sLimit As String
    On Error GoTo errHandler
    
    If Not IsNumber(txtProductLimit.Text) Then
        Cancel = True
        MsgBox "Please enter a valid number for product limit."
        Exit Sub
    End If
    
    dblNum = Var2Dbl(txtProductLimit.Text)
    If dblNum <= 0 Then
        MsgBox "Please enter a positive, non-zero number for product limit."
        Cancel = True
        Exit Sub
    End If
    
    dblProductLimit = dblNum / cGlobalInfo.MassUnits.Multiplier
    Exit Sub
errHandler:
    ErrorIn "frmMain.txtProductLimit_Validate(Cancel)", Cancel
End Sub

Private Sub txtRatio_Validate(Cancel As Boolean)
    Dim dblNum As Double
    On Error GoTo errHandler
    
    If Not IsNumber(txtRatio.Text) Then
        Cancel = True
        MsgBox "Please enter a number from 0 to 100 for Emulsion Percentage"
        Exit Sub
    End If
    
    dblNum = CDbl(txtRatio.Text)
    If dblNum < 0 Or dblNum > 100 Then
        Cancel = True
        MsgBox "Please enter a number from 0 to 100 for Emulsion Percentage"
        Exit Sub
    End If
    
    dblEmulPct = dblNum
    UpdateCapacityDisplay
    Exit Sub
errHandler:
    ErrorIn "frmMain.txtRatio_Validate(Cancel)", Cancel
End Sub

Private Sub chkOnRoad_Click()
    If chkOnRoad.Value = 1 Then
        bIsOnRoad = True
    Else
        bIsOnRoad = False
    End If
    UpdateTagDisplay
End Sub

Private Sub cmdCalculate_Click()
    'Start the solver
    Dim dblMaxProd As Double
    On Error GoTo errHandler
    
    If chkLimitProduct.Value = 0 Then
        dblMaxProd = 0
    Else
        dblMaxProd = dblProductLimit
    End If
    
    Me.MousePointer = vbHourglass
    sngTimeBefore = Timer 'Debugging code
    SolveLoading cActiveTruck, bFindBest, dblMaxProd
    Me.MousePointer = vbDefault
    'MsgBox Timer - sngTimeBefore & " sec"
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdCalculate_Click"
End Sub


Private Sub InitForm()
    Dim i%
    Dim NumAN As Integer
    Dim NumEmul As Integer
    Dim vTemp
    Dim intTemp As Integer
    Dim sTemp As String
    On Error GoTo errHandler

    'Labels
    vTemp = GetFileBaseName(strCurFile)
    lblTruck.Caption = cActiveTruck.Description
    lblFileName.Caption = vTemp
    lblBody.Caption = cActiveTruck.Body.DisplayName
    lblSN.Caption = cActiveTruck.SN
    lblChassis.Caption = cActiveTruck.Chassis.DisplayName
    lblVersion.Caption = cActiveTruck.CreateVersion
    'Mass Units Display
    For Each vTemp In lblMassUnits
        'lblMassUnits(0).Caption = "(" & cGlobalInfo.MassUnits.Display & ")"
        vTemp.Caption = "(" & cGlobalInfo.MassUnits.Display & ")"
    Next
    'Avoid Low Fill of Product Tanks
    mNoLowFill.Checked = cGlobalInfo.AvoidLowFill
    'Tank Configuration
    For i% = 1 To 4
        cboTankConfig(i%).Clear
        cboTankConfig(i%).Visible = False
        cboTankConfig(i%).Enabled = False
        cboTankConfig(i%).AddItem "AN", 0
        cboTankConfig(i%).AddItem "Emulsion", 1
        lblTankConfig(i%).Visible = False
        lblDensity(i%).Visible = False
    Next i%
    vTemp = 0
    For i% = 1 To cActiveTruck.Body.Tanks.Count
        lblDensity(i%).Visible = True
        If cActiveTruck.Body.Tanks(i%).TankType = ttDual Then
            'cboTankConfig(i%).Enabled = true
            lblTankConfig(i%).Caption = cActiveTruck.Body.Tanks(i%).DisplayName
            lblTankConfig(i%).Visible = True
            cboTankConfig(i%).ListIndex = cActiveTruck.Body.Tanks(i%).CurTankUse
            cboTankConfig(i%).Tag = i% 'Tank Number
            cboTankConfig(i%).Visible = True
            lblDensity(i%).Caption = cActiveTruck.Body.Tanks(i%).DensityContents
            vTemp = vTemp + 1
        Else
            lblTankConfig(i%).Caption = cActiveTruck.Body.Tanks(i%).DisplayName
            cboTankConfig(i%).ListIndex = cActiveTruck.Body.Tanks(i%).TankType
            cboTankConfig(i%).Tag = i% 'Tank Number
            lblTankConfig(i%).Visible = True
            cboTankConfig(i%).Visible = True
            lblDensity(i%).Caption = cActiveTruck.Body.Tanks(i%).DensityContents
        End If
    Next i%
    bConfigurable = CBool(vTemp > 0)
    
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

    'Editable Components
    vTemp = -1
    tabComponent.Caption = ""
    tabComponent.RemoveTab 0
    For i% = 1 To cActiveTruck.Components.Count
        If cActiveTruck.Components(i%).ContentsType = ctOther Then
            vTemp = vTemp + 1
            tabComponent.AddTab cActiveTruck.Components(i%).DisplayName
            If cActiveTruck.Components(i%).Capacity.DefaultVolContents <> "" Then
                'Specify by volume
                tabComponent.TabData(vTemp) = "V" & CStr(i%)
                'Set CurVal = Default Val
                cActiveTruck.Components(i%).Capacity.CurVol = cActiveTruck.Components(i%).Capacity.DefaultVolContents
            Else
                'Specify by mass
                tabComponent.TabData(vTemp) = "M" & CStr(i%)
                'Set CurVal = Default Val
                If cActiveTruck.Components(i%).Capacity.DensityContents > 0 Then
                    cActiveTruck.Components(i%).Capacity.CurVol = _
                     cActiveTruck.Components(i%).Capacity.DefaultWtContents / _
                     cActiveTruck.Components(i%).Capacity.DensityContents
                Else
                    cActiveTruck.Components(i%).Capacity.CurVol = 0
                End If
            End If
        End If
    Next i%
    fraComponent.Visible = CBool(vTemp >= 0)
    If vTemp >= 0 Then
        'Activate first tab
        tabComponent.CurrTab = 0
        tabComponent_Click
    End If
    'Emulsion Percentage
    bFindBest = True
    optCalculate(0).Value = True
    For i% = 1 To cActiveTruck.Body.Tanks.Count
        If cActiveTruck.Body.Tanks(i%).TankType = ttAN Then
            NumAN = NumAN + 1
        Else 'Tank is emulsion, or >could< be emulsion
            NumEmul = NumEmul + 1
        End If
    Next i%
    If NumEmul = 0 Or NumAN = 0 Then
        'all AN of all Emul --> no choice
        lblRatio.Visible = False
        txtRatio.Visible = False
        optCalculate(0).Enabled = False
        optCalculate(1).Enabled = False
    Else
        lblRatio.Visible = True
        txtRatio.Visible = True
        optCalculate(0).Enabled = True
        optCalculate(1).Enabled = True
        txtRatio.Text = 30
        dblEmulPct = 30
    End If
    
    'Product Limit
    chkLimitProduct.Value = 0
    txtProductLimit.Enabled = False
    dblProductLimit = 6803.9
    txtProductLimit.Text = Format(6803.9 * cGlobalInfo.MassUnits.Multiplier, "#")
    txtProductLimit.BackColor = BLUE_BKGD
    lblProductLimit.Caption = cGlobalInfo.MassUnits.Display
    UpdateCapacityDisplay
    
    'Update AutoTankFilling button
    intTemp = 0
    For i% = 1 To cActiveTruck.Components.Count
        If cActiveTruck.Components(i%).ContentsType <> ctOther _
         And cActiveTruck.Components(i%).ContentsType <> ctNone Then
            'This is a dependent-style tank
            intTemp = intTemp + 1
        End If
    Next i%
    If intTemp = 0 Then
        'Hide button since it's not applicable to this truck
        lblAutoTankFill.Visible = False
        cmdAutoTankFill.Visible = False
        lin.Visible = False
    Else
        'Show button w/ proper label
        lblAutoTankFill.Visible = True
        cmdAutoTankFill.Visible = True
        lin.Visible = True
        If cGlobalInfo.FillMethod_Additive + cGlobalInfo.FillMethod_Fuel _
         + cGlobalInfo.FillMethod_GasA + cGlobalInfo.FillMethod_GasB _
         + cGlobalInfo.FillMethod_Water > 0 Then
            cmdAutoTankFill.Caption = "Manually"
        Else
            cmdAutoTankFill.Caption = "Automatically"
        End If
    End If
    
    Exit Sub
errHandler:
    ErrorIn "frmMain.InitForm"
End Sub

Private Sub tabComponent_Click()
    Dim a$
    Dim intComp As Integer
    Dim bVol As Boolean
    Dim dblVal As Double
    On Error GoTo errHandler
    
    a$ = tabComponent.TabData(tabComponent.CurrTab)
    bVol = CBool(Mid$(a$, 1, 1) = "V")
    intComp = CInt(Mid$(a$, 2))
    
    'Show/hide controls and display Capacity if appl.
    If bVol Then
        lblCapacity.Visible = True
        lblVolume.Visible = True
        lblCompUnits(1).Visible = True
        lblCompUnits(0).Caption = cGlobalInfo.VolumeUnits.Display
        lblCompUnits(1).Caption = cGlobalInfo.VolumeUnits.Display
        dblVal = Var2Dbl(cActiveTruck.Components(intComp).Capacity.DefaultVolContents) * cGlobalInfo.VolumeUnits.Multiplier
        txtDefault.Text = Format(dblVal, "#0.#")
        dblVal = cActiveTruck.Components(intComp).Capacity.Volume * cGlobalInfo.VolumeUnits.Multiplier
        lblVolume.Caption = Format(dblVal, "#0.#")
        dblVal = cActiveTruck.Components(intComp).Capacity.CurVol * cGlobalInfo.VolumeUnits.Multiplier
        txtCurVal.Text = Format(dblVal, "#0.#")
    Else
        lblCapacity.Visible = False
        lblVolume.Visible = False
        lblCompUnits(1).Visible = False
        lblCompUnits(0).Caption = cGlobalInfo.MassUnits.Display
        dblVal = Var2Dbl(cActiveTruck.Components(intComp).Capacity.DefaultWtContents) * cGlobalInfo.MassUnits.Multiplier
        txtDefault.Text = Format(dblVal, "#0.#")
        dblVal = cActiveTruck.Components(intComp).Capacity.CurVol * cActiveTruck.Components(intComp).Capacity.DensityContents 'kg
        dblVal = dblVal * cGlobalInfo.MassUnits.Multiplier
        txtCurVal.Text = Format(dblVal, "#0.#")
    End If
    Exit Sub
errHandler:
    ErrorIn "frmMain.tabComponent_Click"
End Sub


Private Sub cboTankConfig_Validate(Index As Integer, Cancel As Boolean)
    Call cboTankConfig_Click(Index)
End Sub

Private Sub cboTankConfig_Click(Index As Integer)
    Dim intTank As Integer
    Dim sErr As String
    On Error GoTo errHandler
    
    If cboTankConfig(Index).Tag = "" Then Exit Sub 'Not initialized yet
    intTank = cboTankConfig(Index).Tag 'Tank Number
    cActiveTruck.Body.Tanks(intTank).CurTankUse = cboTankConfig(Index).ListIndex
    sErr = SetProductDensities(cActiveTruck)
    CheckConfigurability
    lblDensity(intTank).Caption = cActiveTruck.Body.Tanks(intTank).DensityContents
    bTruckFileDirty = True
    mSave.Enabled = bTruckFileDirty 'Save only accessible if changes were made
    UpdateCapacityDisplay
    Exit Sub
errHandler:
    ErrorIn "frmMain.cboTankConfig_Click(index)", Index
End Sub

Private Sub optCalculate_Click(Index As Integer)
    bFindBest = CBool(Index = 0)
    CheckConfigurability
    UpdateCapacityDisplay
End Sub


Private Sub CheckConfigurability()
    Dim i%
    Dim NumAN As Integer
    Dim NumEmul As Integer
    Dim sErr As String
    On Error GoTo errHandler
    
    If Not bConfigurable Then Exit Sub 'No configurability, no point continuing...
    
    If bFindBest Then
        'Must allow Emul% editing
        txtRatio.Enabled = True
        For i% = 1 To cActiveTruck.Body.Tanks.Count
            cboTankConfig(i%).Enabled = False
        Next i%
    Else
        'See if a blend is possible with chosen configuration
        For i% = 1 To cActiveTruck.Body.Tanks.Count
            If cActiveTruck.Body.Tanks(i%).TankType = ttDual Then
                cboTankConfig(i%).Enabled = True
                'Make cActiveTruck config matche that indicated by the enabled combobox
                cActiveTruck.Body.Tanks(i%).CurTankUse = cboTankConfig(i%).ListIndex
                lblDensity(i%).Caption = cActiveTruck.Body.Tanks(i%).DensityContents
                sErr = SetProductDensities(cActiveTruck)
            End If
            If cActiveTruck.Body.Tanks(i%).CurTankUse = ttAN Then
                NumAN = NumAN + 1
            Else 'Tank is emulsion
                NumEmul = NumEmul + 1
            End If
        Next i%
        If NumAN > 0 And NumEmul > 0 Then
            'An Blend is possible
            txtRatio.Enabled = True
        Else 'Blend is impossible
            txtRatio.Enabled = False
        End If
    End If
    
    Exit Sub
errHandler:
    ErrorIn "frmMain.CheckConfigurability"
End Sub


Private Sub txtCurVal_Change()
    'Force a validation of any changes
    Dim bDummy As Boolean
    
    Call txtCurVal_Validate(bDummy)
End Sub

Private Sub txtCurVal_Validate(Cancel As Boolean)
    'Check value.  If OK, convert to internal units and save value
    Dim dblMax As Double 'Max value (in approp. eng. units)
    Dim a$
    Dim intComp As Integer 'index of Component being edited
    Dim bVol As Boolean 'TRUE=edit as volume, FALSE=edit as mass
    Dim dblLiters As Double
    Dim dblKG As Double
    Dim dblEditVal As Double
    On Error GoTo errHandler
    
    'Identify the component being edited
    a$ = tabComponent.TabData(tabComponent.CurrTab)
    bVol = CBool(Mid$(a$, 1, 1) = "V")
    intComp = CInt(Mid$(a$, 2))
    
    dblEditVal = Var2Dbl(txtCurVal.Text) 'in user-defined eng. units
    If bVol Then
        dblMax = cActiveTruck.Components(intComp).Capacity.Volume * cGlobalInfo.VolumeUnits.Multiplier
    Else
        dblMax = (cActiveTruck.Components(intComp).Capacity.Volume * cActiveTruck.Components(intComp).Capacity.DensityContents) _
                * cGlobalInfo.MassUnits.Multiplier
    End If
    If dblMax = 0 Then
        'Don't allow Zero as a max
        dblMax = 1.7976931348623E+308 'Huge frick'n number
    End If
    
    'Make user enter a valid preset
    dblEditVal = Var2Dbl(txtCurVal.Text)
    'Check validity of change
    If dblEditVal < 0# Then
        'Give warning, validation fails
        MsgBox "Negative values are not allowed.", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf dblEditVal > dblMax Then
        'Give warning, validation fails
        MsgBox "You cannot enter values greater than " & Format(dblMax, "#0.#"), vbExclamation
        Cancel = True
        Exit Sub
    End If
    
    'Valid new value must have been entered, so save it
    If bVol Then
        'Change Volume directly
        dblLiters = dblEditVal / cGlobalInfo.VolumeUnits.Multiplier
    Else
        'Convert edited mass into volume for internal use
        dblKG = dblEditVal / cGlobalInfo.MassUnits.Multiplier
        dblLiters = dblKG / cActiveTruck.Components(intComp).Capacity.DensityContents
    End If
    cActiveTruck.Components(intComp).Capacity.CurVol = dblLiters
    Exit Sub
errHandler:
    ErrorIn "frmMain.txtCurVal_Validate(Cancel)", Cancel
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
    UpdateTagDisplay
    Exit Sub
errHandler:
    ErrorIn "frmMain.txtTag_Validate(Cancel)", Cancel
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
    ErrorIn "frmMain.UpdateTagDisplay"
End Sub

Private Sub UpdateCapacityDisplay()
    'Since there can only be up to 2 dual-use tanks on any body, there
    ' are a max of 4 possible configurations
    Dim i%
    Dim x%
    Dim NumDualUse As Integer
    Dim NumConfigs As Integer
    Dim dblFullPct As Double
    Dim dblPctDiff As Double
    Dim dblANTot As Double
    Dim dblEmulTot As Double
    Dim dblANTot_Best As Double
    Dim dblEmulTot_Best As Double
    Dim dblVal As Double
    Dim intEmulTanks As Integer
    Dim intANTanks As Integer
    Dim intConfigAtStart As Integer
    On Error GoTo errHandler
    
    intConfigAtStart = CurrentConfigVal(cActiveTruck)
    
    'First find out how many configurations there could be
    For i% = 1 To cActiveTruck.Body.Tanks.Count
        Select Case cActiveTruck.Body.Tanks(i%).TankType
        Case ttDual
            NumDualUse = NumDualUse + 1
        End Select
    Next i%
    
    If NumDualUse = 0 Or (Not bFindBest) Then
        'there's only one possible configuration
        For i% = 1 To cActiveTruck.Body.Tanks.Count
            Select Case cActiveTruck.Body.Tanks(i%).CurTankUse
            Case ttAN
                intANTanks = intANTanks + 1
            Case ttEmulsion
                intEmulTanks = intEmulTanks + 1
            End Select
        Next i%
        lblMaxAN.Visible = (intANTanks <> 0)
        lblMaxEmul.Visible = (intEmulTanks <> 0)
    Else
        lblMaxAN.Visible = True
        lblMaxEmul.Visible = True
    End If
    
    'See which configs are closest to ideal
    NumConfigs = 2 ^ (NumDualUse)
    If Not bFindBest Then NumConfigs = 1 'Only use current config
    For i% = 0 To (NumConfigs - 1)
        If bFindBest Then Call SetTruckConfig(cActiveTruck, i%)
        'sum components assuming full tanks
        dblEmulTot = 0#
        dblANTot = 0#
        For x% = 1 To cActiveTruck.Body.Tanks.Count
            If cActiveTruck.Body.Tanks(x%).CurTankUse = ttEmulsion Then
                'Increment Emulsion total
                dblEmulTot = dblEmulTot + cActiveTruck.Body.Tanks(x%).Volume * cGlobalInfo.DensityEmul
            Else
                'Increment AN total
                dblANTot = dblANTot + cActiveTruck.Body.Tanks(x%).Volume * cGlobalInfo.DensityAN
            End If
        Next x%
        'Calculate %Emul (taking DFO into acct!)
        dblFullPct = Round(100# * dblEmulTot / (dblEmulTot + dblANTot / (1# - dblDFOPct)), 1)
        dblPctDiff = (dblFullPct - dblEmulPct)
        If dblEmulPct = 0 Then
            dblEmulTot = 0
        ElseIf dblEmulPct = 100 Then
            dblANTot = 0
        ElseIf Not bFindBest And (intANTanks = 0 Or intEmulTanks = 0) Then
            'temp over-ride of target %emul since all tanks are same type
        Else
            If dblPctDiff > 0 Then
                'reduce emulsion to meet goal
                dblEmulTot = (dblANTot / (1# - dblDFOPct)) / (100# / dblEmulPct - 1#)
            ElseIf dblPctDiff < 0 Then
                'reduce AN to meet goal
                dblANTot = dblEmulTot * (100# / dblEmulPct - 1#) * (1# - dblDFOPct)
            End If
        End If
        'Check to see if this is best yet
        If (dblANTot + dblEmulTot) > (dblANTot_Best + dblEmulTot_Best) Then
            dblANTot_Best = dblANTot
            dblEmulTot_Best = dblEmulTot
        End If
    Next i%
    
    SetTruckConfig cActiveTruck, intConfigAtStart
    dblVal = dblANTot_Best * cGlobalInfo.MassUnits.Multiplier
    lblMaxAN.Caption = "Maximum Capacity AN = " & Format(dblVal, "#,##0") & cGlobalInfo.MassUnits.Display
    dblVal = dblEmulTot_Best * cGlobalInfo.MassUnits.Multiplier
    lblMaxEmul.Caption = "Maximum Capacity Emulsion = " & Format(dblVal, "#,##0") & cGlobalInfo.MassUnits.Display
    Exit Sub
errHandler:
    ErrorIn "frmMain.UpdateCapacityDisplay"
End Sub


Private Sub tabTagDisplay_Click()
    UpdateTagDisplay
End Sub

