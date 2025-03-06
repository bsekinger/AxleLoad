VERSION 5.00
Begin VB.Form frmComponentBuilder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Component Builder"
   ClientHeight    =   6972
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   8892
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6972
   ScaleWidth      =   8892
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraContents 
      Caption         =   "Contents"
      Height          =   5892
      Left            =   3600
      TabIndex        =   19
      Top             =   960
      Width           =   5172
      Begin VB.Frame fraEmulsion 
         Height          =   2052
         Left            =   2640
         TabIndex        =   45
         Top             =   3720
         Width           =   2412
         Begin VB.CheckBox chkEmulsion 
            Caption         =   "Emulsion Relationship"
            Height          =   372
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   1932
         End
         Begin VB.TextBox txtMultiplier 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   876
            Width           =   1332
         End
         Begin VB.TextBox txtOffset 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   46
            Top             =   1596
            Width           =   1332
         End
         Begin VB.Label lblMultiplier 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier"
            Height          =   280
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   1332
         End
         Begin VB.Label lblOffset 
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
            Height          =   280
            Index           =   3
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   1332
         End
         Begin VB.Label lblMassUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass"
            Height          =   280
            Index           =   3
            Left            =   1560
            TabIndex        =   49
            Top             =   1680
            Width           =   732
         End
      End
      Begin VB.Frame fraAN 
         Height          =   2052
         Left            =   120
         TabIndex        =   38
         Top             =   3720
         Width           =   2412
         Begin VB.CheckBox chkAN 
            Caption         =   "AN Relationship"
            Height          =   372
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   1932
         End
         Begin VB.TextBox txtMultiplier 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   876
            Width           =   1332
         End
         Begin VB.TextBox txtOffset 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   1596
            Width           =   1332
         End
         Begin VB.Label lblMultiplier 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier"
            Height          =   276
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   1332
         End
         Begin VB.Label lblOffset 
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
            Height          =   276
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   1332
         End
         Begin VB.Label lblMassUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "Mass"
            Height          =   276
            Index           =   2
            Left            =   1560
            TabIndex        =   41
            Top             =   1680
            Width           =   732
         End
      End
      Begin VB.CommandButton cmdContentCG 
         Caption         =   "Edit Coefficients"
         Height          =   372
         Left            =   2520
         TabIndex        =   37
         Top             =   3120
         Width           =   1692
      End
      Begin VB.CommandButton cmdStickLength 
         Caption         =   "Edit Coefficients"
         Height          =   372
         Left            =   2520
         TabIndex        =   34
         Top             =   1440
         Width           =   1572
      End
      Begin VB.TextBox txtVolume 
         Height          =   315
         Left            =   2520
         TabIndex        =   31
         Top             =   636
         Width           =   1452
      End
      Begin VB.CheckBox chkUsesSightGauge 
         Caption         =   "Uses Sight Gauge"
         Height          =   252
         Left            =   2520
         TabIndex        =   30
         Top             =   2160
         Width           =   2172
      End
      Begin VB.TextBox txtDensityContents 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   3156
         Width           =   1212
      End
      Begin VB.TextBox txtDefaultVolContents 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   2316
         Width           =   1212
      End
      Begin VB.TextBox txtDefaultWtContents 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1476
         Width           =   1212
      End
      Begin VB.ComboBox cboContentsType 
         Height          =   288
         Left            =   120
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   636
         Width           =   1812
      End
      Begin VB.Label lblContentCG 
         Caption         =   "Content CG Equation"
         Height          =   276
         Left            =   2520
         TabIndex        =   36
         Top             =   2880
         Width           =   1812
      End
      Begin VB.Label lblStickLength 
         Caption         =   "Stick Length Equation"
         Height          =   276
         Left            =   2520
         TabIndex        =   35
         Top             =   1200
         Width           =   1692
      End
      Begin VB.Label lblVolumeUnits 
         Caption         =   "VolUnits"
         Height          =   276
         Index           =   1
         Left            =   4080
         TabIndex        =   33
         Top             =   756
         Width           =   732
      End
      Begin VB.Label lblVolume 
         Caption         =   "Capacity (Volume)"
         Height          =   276
         Left            =   2520
         TabIndex        =   32
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label lblDensityContents 
         Caption         =   "Specific Gravity"
         Height          =   276
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   1212
      End
      Begin VB.Label lblVolumeUnits 
         Caption         =   "VolUnits"
         Height          =   276
         Index           =   0
         Left            =   1440
         TabIndex        =   27
         Top             =   2436
         Width           =   732
      End
      Begin VB.Label lblDefaultVolContents 
         Caption         =   "Default Volume"
         Height          =   276
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1212
      End
      Begin VB.Label lblMassUnits 
         Caption         =   "Mass"
         Height          =   276
         Index           =   1
         Left            =   1440
         TabIndex        =   24
         Top             =   1596
         Width           =   732
      End
      Begin VB.Label lblDefaultWtContents 
         Caption         =   "Default Weight"
         Height          =   276
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label lblContentsType 
         Caption         =   "Type"
         Height          =   276
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2172
      End
   End
   Begin VB.Frame fraEmpty 
      Caption         =   "Empty"
      Height          =   1812
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   2292
      Begin VB.TextBox txtEmptyCG 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1356
         Width           =   1212
      End
      Begin VB.TextBox txtEmptyWeight 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   636
         Width           =   1212
      End
      Begin VB.Label lblMassUnits 
         Caption         =   "Mass"
         Height          =   276
         Index           =   0
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   732
      End
      Begin VB.Label lblDistanceUnits 
         Caption         =   "DistUnits"
         Height          =   276
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label lblEmptyCG 
         Caption         =   "CG Location"
         Height          =   276
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label lblEmptyWeight 
         Caption         =   "Weight"
         Height          =   276
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.TextBox txtOffset 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1332
   End
   Begin VB.ComboBox cboLocationReference 
      Height          =   288
      Left            =   120
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1800
      Width           =   1812
   End
   Begin VB.TextBox txtDisplayName 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3252
   End
   Begin VB.TextBox txtFullName 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   6732
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   492
      Left            =   2280
      TabIndex        =   1
      Top             =   6360
      Width           =   1092
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   1092
   End
   Begin VB.Label lblDistanceUnits 
      Caption         =   "DistUnits"
      Height          =   276
      Index           =   0
      Left            =   1560
      TabIndex        =   11
      Top             =   2640
      Width           =   732
   End
   Begin VB.Label lblFullName 
      Caption         =   "Full Name"
      Height          =   276
      Left            =   120
      TabIndex        =   3
      Top             =   84
      Width           =   1812
   End
   Begin VB.Label lblDisplayName 
      Caption         =   "Display Name"
      Height          =   276
      Left            =   120
      TabIndex        =   5
      Top             =   804
      Width           =   2052
   End
   Begin VB.Label lblLocationReference 
      Caption         =   "Locating Reference Point"
      Height          =   276
      Left            =   120
      TabIndex        =   7
      Top             =   1524
      Width           =   1812
   End
   Begin VB.Label lblOffset 
      Caption         =   "Offset from Reference Point"
      Height          =   276
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2244
      Width           =   2052
   End
End
Attribute VB_Name = "frmComponentBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim vLabel
    
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


End Sub

