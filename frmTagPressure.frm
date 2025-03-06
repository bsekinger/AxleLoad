VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmTagPressure 
   Caption         =   "Calibrate Tag Pressure Factors"
   ClientHeight    =   4044
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6984
   Icon            =   "frmTagPressure.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4044
   ScaleWidth      =   6984
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   852
      Left            =   5520
      TabIndex        =   2
      Top             =   3000
      Width           =   1332
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   852
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   1332
   End
   Begin C1SizerLibCtl.C1Tab tabTagDisplay 
      Height          =   1332
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   3372
      _cx             =   5948
      _cy             =   2350
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
         Height          =   1044
         Left            =   12
         ScaleHeight     =   1044
         ScaleWidth      =   3348
         TabIndex        =   4
         Top             =   276
         Width           =   3348
         Begin VB.TextBox txtWeight 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   516
            Width           =   1212
         End
         Begin VB.TextBox txtPressure 
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   552
            Width           =   1212
         End
         Begin VB.Label lblLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Scale Weight"
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
            TabIndex        =   10
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label lblMassUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "kg"
            Height          =   276
            Index           =   0
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   372
         End
         Begin VB.Label lblPressure 
            BackStyle       =   0  'Transparent
            Caption         =   "Pressure"
            Height          =   276
            Left            =   1800
            TabIndex        =   8
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label lblEqual 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "="
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
            Index           =   1
            Left            =   1320
            TabIndex        =   7
            Top             =   600
            Width           =   492
         End
      End
   End
   Begin VB.Label lblDuh 
      BackStyle       =   0  'Transparent
      Caption         =   "Nothing to adjust -- this truck has no tag axles."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   276
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   3852
   End
   Begin VB.Label lblWarning 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTagPressure.frx":030A
      ForeColor       =   &H80000008&
      Height          =   2052
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6732
   End
End
Attribute VB_Name = "frmTagPressure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cCurTruck As clsTruck 'used in this form as a link to the active truck
Private cTags As clsTags 'temp copy of Tags collection

Public Sub Edit(cTruck As clsTruck)
    Set cCurTruck = cTruck
    Me.Show
End Sub


Private Sub cmdApply_Click()
    SaveChanges
End Sub

Private Sub Form_Load()
    Dim i%
    Dim dblDefault As Double
    Dim dblVal As Double
    Dim cTag As clsTag
    
    cmdApply.Enabled = False
    tabTagDisplay.Visible = False
    
    'Tag options
    If cCurTruck.Chassis.Tags.Count > 0 Then
        'One or more tags
        lblDuh.Visible = False
        cmdApply.Enabled = False
        tabTagDisplay.Visible = True
        tabTagDisplay.Caption = ""
        tabTagDisplay.RemoveTab 0
        'Add tags to display and (if necessary) a temporary class collection
        Set cTags = New clsTags
        For i% = 1 To cCurTruck.Chassis.Tags.Count
            If cCurTruck.Chassis.Tags(i%).Location < cCurTruck.Chassis.WB Then
                tabTagDisplay.AddTab i% & " (Pusher)"
            Else
                tabTagDisplay.AddTab i% & " (Tag)"
            End If
            tabTagDisplay.TabData(i% - 1) = 2000# 'default for mass/wt = 2000
            'Add tag to temp collection and fill tag with data
            Set cTag = New clsTag
            With cTag
                .DownwardForce = cCurTruck.Chassis.Tags(i%).DownwardForce
                .ForceToPressure = cCurTruck.Chassis.Tags(i%).ForceToPressure
                .Location = cCurTruck.Chassis.Tags(i%).Location
                .Weight = cCurTruck.Chassis.Tags(i%).Weight
                .WtLimit = cCurTruck.Chassis.Tags(i%).WtLimit
            End With
            cTags.Add cTag
        Next
        'Set the first tab as the active tab
        tabTagDisplay.CurrTab = 0
        tabTagDisplay_Click
    Else 'No Tags
        lblDuh.Visible = True
        tabTagDisplay.Visible = False
        cmdApply.Enabled = False
        Exit Sub
    End If
    tabTagDisplay.CurrTab = 0
    tabTagDisplay_Click
    Set cTag = Nothing
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cCurTruck = Nothing
    Set cTags = Nothing
End Sub


Private Sub SaveChanges()
    Set cCurTruck.Chassis.Tags = New clsTags
    Set cCurTruck.Chassis.Tags = cTags
    Unload Me
End Sub


Private Sub tabTagDisplay_Click()
    'User chose a tab, update it
    Dim i%
    Dim dblDefault As Double
    Dim dblVal As Double
    
    If lblDuh.Visible Then Exit Sub 'no Tags to deal with
    
    dblDefault = tabTagDisplay.TabData(tabTagDisplay.CurrTab)
    txtWeight.Text = dblDefault
    i% = tabTagDisplay.CurrTab + 1
    dblVal = dblDefault / cGlobalInfo.MassUnits.Multiplier * cTags(i%).ForceToPressure
    txtPressure.Text = Format(dblVal, "#0")

    cmdApply.Enabled = True
End Sub


Private Function CalcPressRelationship()
    Dim i%
    
    i% = tabTagDisplay.CurrTab + 1
    cTags(i%).ForceToPressure = CDbl(txtPressure.Text) / (CDbl(txtWeight.Text) / cGlobalInfo.MassUnits.Multiplier)
End Function


'=================================================
Private Sub txtWeight_Validate(Cancel As Boolean)
    Dim dblNum As Double
    
    If Trim$(txtWeight.Text) = "" Then
        'will treat this as zero
    ElseIf Not IsNumber(txtWeight.Text) Then
        Cancel = True
        MsgBox "Please enter a valid number for Scale Weight"
        Exit Sub
    End If
    
    If Trim$(txtWeight.Text) = "" Then
        dblNum = 0
    Else
        dblNum = Var2Dbl(txtWeight.Text)
    End If
    
    If dblNum < 0 Then
        MsgBox "Please enter positive number for Scale Weight"
        Cancel = True
        Exit Sub
    End If
    tabTagDisplay.TabData(tabTagDisplay.CurrTab) = dblNum
    CalcPressRelationship
End Sub


'=================================================
Private Sub txtPressure_Validate(Cancel As Boolean)
    Dim dblNum As Double
    
    If Trim$(txtPressure.Text) = "" Then
        'will treat this as zero
    ElseIf Not IsNumber(txtPressure.Text) Then
        Cancel = True
        MsgBox "Please enter a valid number for Air Pressure"
        Exit Sub
    End If
    
    If Trim$(txtPressure.Text) = "" Then
        dblNum = 0
    Else
        dblNum = Var2Dbl(txtPressure.Text)
    End If
    
    If dblNum < 0 Then
        MsgBox "Please enter positive number for Air Pressure"
        Cancel = True
        Exit Sub
    End If
    CalcPressRelationship
End Sub

