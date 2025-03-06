VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{9439E91F-1836-11D3-8E38-444553540000}#3.0#0"; "dxfreader.ocx"
Begin VB.Form frmDXF 
   Caption         =   "Form1"
   ClientHeight    =   5952
   ClientLeft      =   132
   ClientTop       =   816
   ClientWidth     =   9096
   LinkTopic       =   "Form1"
   ScaleHeight     =   5952
   ScaleWidth      =   9096
   StartUpPosition =   3  'Windows Default
   Begin DXFREADERlib.DXFReader DXFReader2 
      Height          =   492
      Left            =   8160
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   868
      RegistrationCode=   "61X94OO862244165"
      PlotMode        =   4
      PictureScaleMode=   3
      PlotRendering   =   0
      PlotRotation    =   0
      PlotPenWidth    =   1
      FileStatus      =   0
      MinX            =   0
      MinY            =   0
      MaxX            =   100
      MaxY            =   100
      ScaleX          =   1
      ScaleY          =   1
      TranslationX    =   0
      TranslationY    =   0
      PlotScale       =   1
      BaseX           =   0
      BaseY           =   0
      PictureBaseX    =   0
      PictureBaseY    =   0
      PictureWidth    =   0
      PictureHeight   =   0
      PictureScaleX   =   1
      PictureScaleY   =   1
      RotationAngle   =   0
      Version         =   "1.59.0051"
      ZoomInOutPercent=   50
      AutoRedraw      =   -1  'True
      PaletteCaption  =   "Select Color"
      PaletteCancelButtonText=   "Cancel"
      PaletteOkButtonText=   "Ok"
      MouseIcon       =   "frmDXF.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrResize 
      Left            =   8520
      Top             =   2880
   End
   Begin DXFREADERlib.DXFReader DXFReader1 
      Height          =   5292
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7812
      _ExtentX        =   13780
      _ExtentY        =   9335
      RegistrationCode=   "61X94OO862244165"
      PlotMode        =   4
      PictureScaleMode=   3
      PlotRendering   =   0
      PlotRotation    =   0
      PlotPenWidth    =   1
      FileStatus      =   0
      MinX            =   0
      MinY            =   0
      MaxX            =   119.972
      MaxY            =   59.972
      ScaleX          =   1
      ScaleY          =   1
      TranslationX    =   0
      TranslationY    =   0
      PlotScale       =   1
      BaseX           =   0
      BaseY           =   0
      PictureBaseX    =   0
      PictureBaseY    =   0
      PictureWidth    =   0
      PictureHeight   =   0
      PictureScaleX   =   1
      PictureScaleY   =   1
      RotationAngle   =   0
      Version         =   "1.59.0051"
      ZoomInOutPercent=   50
      AutoRedraw      =   -1  'True
      PaletteCaption  =   "Select Color"
      PaletteCancelButtonText=   "Cancel"
      PaletteOkButtonText=   "Ok"
      MouseIcon       =   "frmDXF.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CurrentTextStyle=   "STANDARD"
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   8520
      Top             =   3360
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mSaveAs 
         Caption         =   "Sav&e As"
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mOtherSide 
      Caption         =   "&Other Side"
   End
End
Attribute VB_Name = "frmDXF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cCurTruck As clsTruck

Private bCurbSide As Boolean
Private AspectRatio As Double

Public Sub ShowDrawing(cTruck As clsTruck)
    Set cCurTruck = cTruck
    bOKToResize = False
    mOtherSide.Caption = ""
    Me.Show vbModal, frmMain
    
    Set cCurTruck = Nothing
End Sub



Private Sub Form_Activate()
    
    bOKToResize = True
    Me.WindowState = vbMaximized
    'Form_Resize
    tmrResize.Enabled = False
    Me.MousePointer = vbHourglass
    mOtherSide.Caption = "--wait--"
    mOtherSide.Enabled = False
    DoEvents
    RenderTruck cCurTruck, DXFReader1, bCurbSide

    Me.MousePointer = vbDefault
    
    If bCurbSide Then
        mOtherSide.Caption = "Curb Side"
    Else
        mOtherSide.Caption = "Street Side"
    End If
    mOtherSide.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bOKToResize = False
    'DXFReader1.Clear
End Sub

Private Sub Form_Resize()
    Dim WindowBorderY As Long
    Dim WindowBorderX As Long
    Dim CapHt As Long
    Dim MnuHt As Long
    
    If Not bOKToResize Then Exit Sub
    WindowBorderX = GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX 'twips
    WindowBorderY = GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY 'twips
    CapHt = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
    MnuHt = GetSystemMetrics(SM_CYMENUSIZE) * Screen.TwipsPerPixelY
    
    With Me
        DXFReader1.Move 100, _
                        100, _
                        (.Width - 2 * WindowBorderX) - 200, _
                        (.Height - 2 * WindowBorderY - CapHt - MnuHt) - 200
    End With
    
    DoEvents
    
    'Only Zoom Extents when mouse is let go
    tmrResize.Enabled = False
    tmrResize.Interval = 750
    tmrResize.Enabled = True

End Sub

Private Sub DXFReader1_Error(ErrorCode As Integer, ErrorString As String)
    If ErrorCode <> 0 Then
        MsgBox "DXFReader Err" & ErrorCode & ":" & ErrorString
    End If
End Sub


Private Sub mOtherSide_Click()
    bCurbSide = Not bCurbSide
    Me.MousePointer = vbHourglass
    DoEvents
    DoEvents
    DoEvents
    Me.MousePointer = vbHourglass
    DoEvents
    DoEvents
    
    mOtherSide.Caption = "--wait--"
    mOtherSide.Enabled = False
    
    DXFReader1.Clear
    Form_Resize
    RenderTruck cCurTruck, DXFReader1, bCurbSide
    
    If bCurbSide Then
        mOtherSide.Caption = "Curb Side"
    Else
        mOtherSide.Caption = "Street Side"
    End If
    mOtherSide.Enabled = True
    
End Sub


Private Sub mPrint_Click()
    Dim fl$
    
    fl$ = AddBackslash(App.Path) & "~temp.wmf"
    'fill a second control so that we can save the WMF
    AspectRatio = (DXFReader1.MaxX - DXFReader1.MinX) / (DXFReader1.MaxY - DXFReader1.MinY)
    With DXFReader2
        .Height = 12000
        .Width = .Height * AspectRatio
    End With
    DXFReader2.FileName = DXFReader1.FileName
    DXFReader2.SaveWMF fl$
    'Now Call the Print Report
    frmReport.PrintDXF cCurTruck, fl$, AspectRatio
End Sub


Private Sub mSaveAs_Click()
    'Save the modified truck to a new file
    Dim sMsg As String
    Dim strNewFile As String
    Dim sFile As String
    On Error GoTo errHandler
    
    If IsNumber(cCurTruck.SN) Then
        'User (properly) eneterd a number or .SN property
        sFile = "Truck_SN" & Trim$(cCurTruck.SN)
    Else
        If Len(cCurTruck.SN) > 2 Then
            If Mid$(cCurTruck.SN, 1, 1) = "q" Then
                sFile = "Truck_Q" & Mid$(cCurTruck.SN, 2)
            ElseIf UCase$(Mid$(cCurTruck.SN, 1, 2)) = "SN" Then
                sFile = "Truck_SN" & Mid$(cCurTruck.SN, 3)
            Else
                sFile = "Truck_" & Trim$(cCurTruck.SN)
            End If
        Else
            sFile = "Truck_" & Trim$(cCurTruck.SN)
        End If
    End If
    If bCurbSide Then
        sFile = sFile & "_CurbView.dxf"
    Else
        sFile = sFile & "_StreetView.dxf"
    End If
    
    With frmMain.dlgFile
        strNewFile = ""
        .FileName = sFile 'strCurFile
        .Filter = "DXF File|*.dxf"
        .DialogTitle = "Save Truck Drawing"
        .CancelError = True
        .InitDir = cGlobalInfo.TruckFolder
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        On Error Resume Next
        .ShowSave 'Show the SaveFile dialog
        'If User selected a file, open it
        strNewFile = Trim$(.FileName)
        If Err <> 0 Or strNewFile = "" Then
            'User canceled
            Exit Sub
        End If
        On Error GoTo errHandler
        strNewFile = Trim$(.FileName)
    End With
    
    DXFReader1.WriteDXF strNewFile
    Exit Sub
errHandler:
    ErrorIn "frmDXF.mSaveAs_Click"
End Sub


Private Sub tmrResize_Timer()
    'Trigger a redraw 750 msec after user stops resizing form
    Me.MousePointer = vbHourglass
    DoEvents
    DXFReader1.ZoomExtents
    Me.MousePointer = vbDefault
    tmrResize.Interval = 0
End Sub
