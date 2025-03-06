VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{D3F92121-EFAA-4B5C-B91B-3D6A8FFD1477}#1.0#0"; "vsdraw8.ocx"
Object = "{9439E91F-1836-11D3-8E38-444553540000}#3.0#0"; "DXFReader.ocx"
Begin VB.Form frmReport 
   Caption         =   "frmReport"
   ClientHeight    =   7500
   ClientLeft      =   132
   ClientTop       =   816
   ClientWidth     =   7956
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7956
   StartUpPosition =   3  'Windows Default
   Begin DXFREADERlib.DXFReader DXFReader2 
      Height          =   612
      Left            =   6600
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   1080
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
      Version         =   "1.59.0053"
      ZoomInOutPercent=   50
      AutoRedraw      =   -1  'True
      PaletteCaption  =   "Select Color"
      PaletteCancelButtonText=   "Cancel"
      PaletteOkButtonText=   "Ok"
      MouseIcon       =   "frmReport.frx":0000
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
   Begin DXFREADERlib.DXFReader DXFReader1 
      Height          =   612
      Left            =   6600
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   1080
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
      Version         =   "1.59.0053"
      ZoomInOutPercent=   50
      AutoRedraw      =   -1  'True
      PaletteCaption  =   "Select Color"
      PaletteCancelButtonText=   "Cancel"
      PaletteOkButtonText=   "Ok"
      MouseIcon       =   "frmReport.frx":001C
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
   Begin VSPrinter8LibCtl.VSPrinter vsPrinter 
      Height          =   7452
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6372
      _cx             =   11239
      _cy             =   13144
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   720
      MarginTop       =   1440
      MarginRight     =   720
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   42.3484848484849
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VSDraw8LibCtl.VSDraw vsDraw 
         Height          =   972
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   4692
         _cx             =   8276
         _cy             =   1714
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   75
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
   End
   Begin VB.Image imgLogo 
      Height          =   324
      Left            =   6600
      Picture         =   "frmReport.frx":0038
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   924
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mSaveAs 
         Caption         =   "Save DXF"
      End
   End
   Begin VB.Menu mOtherSide 
      Caption         =   "Show Other Side"
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cCurTruck As clsTruck
Private bCurbSide As Boolean
Private AspectRatio As Double

Public Sub PrintPreview(cTruck As clsTruck)
    'This sub is called to create a loading report.  It shows a simplified picture
    ' of the truck along with loading information.
    Dim i%
    Dim a$, b$
    Dim dblTotalProduct As Double '(kg) AN + Emul + Fuel
    Dim dblLoad As Double
    Dim dblGVW As Double
    Dim dblTemp As Double
    Dim sMsg As String
    Dim sWarnings() As String
    Dim curY
    Dim vLineHt
    Dim xCol(8) 'X-Position of column(i)
    Dim cAxles As clsAxleGroups
    
    Dim cEmulTanks As New clsTanks
    Dim cANTanks As New clsTanks
    
    Set cANTanks = ANTanks(cTruck)
    Set cEmulTanks = EmulsionTanks(cTruck)
    
    mFile.Visible = False
    mOtherSide.Visible = False
    Me.Show
    
    UpdateDrawing vsDraw, cTruck

    With vsPrinter
        .StartDoc
        
        .MarginLeft = "0.5in"
        .MarginRight = "0.5in"
        .MarginTop = "0.5in"
        .MarginBottom = "0.5in"
        
        .FontName = "Tahoma"
        .FontSize = 12
        .CalcParagraph = "H"
        vLineHt = .TextHei
        
        .FontSize = 28
        .CurrentY = ".5in"
        .TextAlign = taCenterTop
        .Paragraph = cTruck.Description
        
        .FontSize = 12
        .Paragraph = "Unibody Loading Analysis"
        
        .DrawPicture vsDraw.Picture, "2.1in", .CurrentY, "4.3in", "1.4in"  ', vppaCenterTop
        .CurrentY = .Y2
        
        'Percent Emulsion (if applicable)
        curY = .CurrentY
        If cEmulTanks.Count > 0 And cANTanks.Count > 0 Then
            If (cTruck.Body.MassANTotal / (1# - dblDFOPct) + cTruck.Body.MassEmulTotal) = 0 Then
                'Division by zero
                a$ = "0% Emulsion"
            Else
                dblTemp = 100# * cTruck.Body.MassEmulTotal / (cTruck.Body.MassANTotal / (1# - dblDFOPct) + cTruck.Body.MassEmulTotal)
                a$ = Format(dblTemp, "#0") & "% Emulsion"
            End If
            .Paragraph = a$
            curY = .CurrentY
        End If
        .CurrentY = curY + vLineHt 'place cursor incase %Emul n/a

        'Headers =================================================
        .FontSize = 21
        .TextAlign = taLeftTop
        .Paragraph = "TankFill"
        
        .CurrentY = .CurrentY + vLineHt / 2
        .FontSize = 12
        .FontBold = True
        .CurrentX = "1in"
        xCol(1) = .CurrentX
        .Text = "Tank"
        
        .CurrentX = "1.6in"
        xCol(2) = .CurrentX
        .Text = "Contents"
        
        .CurrentX = "2.6in"
        xCol(3) = .CurrentX
        .Text = "Density"
        
        .CurrentX = "3.45in"
        xCol(4) = .CurrentX
        .Text = "Contents"
        .FontBold = False
        .Text = " (" & cGlobalInfo.MassUnits.Display & ")"
        .FontBold = True
        
        .CurrentX = "4.75in"
        xCol(5) = .CurrentX
        .Text = "%Full"
                
        .CurrentX = "5.5in"
        xCol(6) = .CurrentX
        .Text = "Fill Level"
        curY = .CurrentY
        
        'Line =================================================
        .CurrentY = .CurrentY + 1.25 * vLineHt
        .DrawLine "1in", .CurrentY, .PageWidth - .MarginRight, .CurrentY
                
        'Tanks =================================================
        .FontBold = False
        .CurrentY = curY + 1.5 * vLineHt
        For i% = 1 To cTruck.Body.Tanks.Count
            'Tank
            .CurrentX = xCol(1)
            a$ = Chr$(64 + i%)
            .Text = a$
            'Contents
            .CurrentX = xCol(2)
            If cTruck.Body.Tanks(i%).CurTankUse = ttAN Then
                .Text = "AN Prill"
            Else
                .Text = "Emulsion"
            End If
            'Density
            .CurrentX = xCol(3)
            .Text = Format(cTruck.Body.Tanks(i%).DensityContents, "0.00")
            'Contents
            .CurrentX = xCol(4)
            dblTemp = cTruck.Body.Tanks(i%).CurVol * cTruck.Body.Tanks(i%).DensityContents
            dblTotalProduct = dblTotalProduct + dblTemp
            dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblTemp, "#,##0 ") '& cGlobalInfo.MassUnits.Display
            '%Full
            .CurrentX = xCol(5)
            dblTemp = cTruck.Body.Tanks(i%).CurVol / cTruck.Body.Tanks(i%).Volume * 100#
            .Text = Format(dblTemp, "#0") & "%"
            'Fill
            .CurrentX = xCol(6)
            dblTemp = cTruck.Body.Tanks(i%).CurStkHt * cGlobalInfo.DistanceUnits.Multiplier
            If cGlobalInfo.DistanceUnits.Display = "in" Then
                a$ = "#0 "
            Else
                a$ = "#0.## "
            End If
            .Text = Format(dblTemp, a$) & cGlobalInfo.DistanceUnits.Display & " from Top"
            .CurrentY = .CurrentY + 1.5 * vLineHt
        Next i%
        
        'Fuel or Additive (if applic) ============================================
        For i% = 1 To cTruck.Components.Count
            If cTruck.Components(i%).ContentsType = ctFuel Or _
            cTruck.Components(i%).ContentsType = ctAdditive Then
                'Tank
                .CurrentX = xCol(1)
                .Text = cTruck.Components(i%).DisplayName
                'Density
                .CurrentX = xCol(3)
                .Text = Format(cTruck.Components(i%).Capacity.DensityContents, "0.00")
                'Contents
                .CurrentX = xCol(4)
                dblTemp = cTruck.Components(i%).Capacity.CurVol * cTruck.Components(i%).Capacity.DensityContents
                dblTotalProduct = dblTotalProduct + dblTemp
                dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
                .Text = Format(dblTemp, "#,##0 ") '& cGlobalInfo.MassUnits.Display
                '%Full
                .CurrentX = xCol(5)
                dblTemp = cTruck.Components(i%).Capacity.CurVol / cTruck.Components(i%).Capacity.Volume * 100#
                .Text = Format(dblTemp, "#0") & "%"
                'Fill
                .CurrentX = xCol(6)
                If cTruck.Components(i%).Capacity.UsesSightGauge Then
                    'Show as sight gauge (percent)
                    .Text = Format(dblTemp, "#0") & "%"
                Else
                    'Show as fill from top
                    dblTemp = cTruck.Components(i%).Capacity.CurStkHt * cGlobalInfo.DistanceUnits.Multiplier
                    If cGlobalInfo.DistanceUnits.Display = "in" Then
                        a$ = "#0 "
                    Else
                        a$ = "#0.## "
                    End If
                    .Text = Format(dblTemp, a$) & cGlobalInfo.DistanceUnits.Display & " from Top"
                End If
                .CurrentY = .CurrentY + 1.5 * vLineHt
            End If
        Next i%
        
        'Product Total ==============================================
        .FontItalic = True
        .FontBold = True
        .CurrentX = xCol(1)
        .Text = "Total Product"
        'Contents
        .CurrentX = xCol(4)
        dblTemp = dblTotalProduct * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .FontItalic = False
        .FontBold = False
        
        'Non-fuel components ==============================================
        .CurrentY = .CurrentY + 2 * vLineHt
        For i% = 1 To cTruck.Components.Count
            If cTruck.Components(i%).ContentsType <> ctFuel And _
            cTruck.Components(i%).ContentsType <> ctAdditive And _
            cTruck.Components(i%).ContentsType <> ctNone Then
                'Tank
                .CurrentX = xCol(1)
                .Text = cTruck.Components(i%).DisplayName
                'Density
                If cTruck.Components(i%).Capacity.DefaultWtContents <> "" Then
                    'Mass-based, don't show density
                Else
                    'show density
                    .CurrentX = xCol(3)
                    .Text = Format(cTruck.Components(i%).Capacity.DensityContents, "0.00")
                End If
                'Contents
                .CurrentX = xCol(4)
                dblTemp = cTruck.Components(i%).Capacity.CurVol * cTruck.Components(i%).Capacity.DensityContents
                dblTotalProduct = dblTotalProduct + dblTemp
                dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
                .Text = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
                '%Full
                .CurrentX = xCol(5)
                If cTruck.Components(i%).ContentsType = ctOther Then
                    'NOT gassing, water, fuel, or additive
                    If cTruck.Components(i%).Capacity.DefaultWtContents <> "" Then
                        'Mass-based, no volume to show
                    Else
                        dblTemp = cTruck.Components(i%).Capacity.CurVol / cTruck.Components(i%).Capacity.Volume * 100#
                        .Text = Format(dblTemp, "#0") & "%"
                    End If
                Else
                    'This should be a 'normal' volume-type tank
                    dblTemp = cTruck.Components(i%).Capacity.CurVol / cTruck.Components(i%).Capacity.Volume * 100#
                    .Text = Format(dblTemp, "#0") & "%"
                End If
                'Fill
                .CurrentX = xCol(6)
                If cTruck.Components(i%).Capacity.DefaultWtContents <> "" Then
                    'Contents indicated by weight which is already shown
                ElseIf cTruck.Components(i%).Capacity.DefaultVolContents <> "" Then
                    'Show contents by volume
                    dblTemp = cTruck.Components(i%).Capacity.CurVol * cGlobalInfo.VolumeUnits.Multiplier
                    .Text = Format(dblTemp, "#,##0 ") & cGlobalInfo.VolumeUnits.Display
                Else 'show by sightgauge or dist from top
                    If cTruck.Components(i%).Capacity.UsesSightGauge Then
                        'Show as sight gauge (percent)
                        .Text = Format(dblTemp, "#0") & "%"
                    Else
                        'Show as fill from top
                        dblTemp = cTruck.Components(i%).Capacity.CurStkHt * cGlobalInfo.DistanceUnits.Multiplier
                        If cGlobalInfo.DistanceUnits.Display = "in" Then
                            a$ = "#0 "
                        Else
                            a$ = "#0.## "
                        End If
                        .Text = Format(dblTemp, a$) & cGlobalInfo.DistanceUnits.Display & " from Top"
                    End If
                End If
                .CurrentY = .CurrentY + 1.5 * vLineHt
            End If
        Next i%
        
        'AxleLoading Headers ==============================================
        .CurrentY = .CurrentY + 2 * vLineHt
        .CurrentX = .MarginLeft
        .FontSize = 21
        .Text = "Axle Loading"
        .FontSize = 12
        .Text = " (" & cGlobalInfo.MassUnits.Display & ")"

        
        .CurrentY = .CurrentY + vLineHt / 2
        .FontSize = 10
        .FontBold = True
        .CurrentX = "1in"
        xCol(1) = .CurrentX
        
        .CurrentX = "2.9in"
        xCol(2) = .CurrentX
        
        .CurrentX = "3.65in"
        xCol(3) = .CurrentX
        
        .CurrentX = "4.4in"
        xCol(4) = .CurrentX
        
        .CurrentX = "5.15in"
        xCol(5) = .CurrentX
        
        .CurrentX = "5.9in"
        xCol(6) = .CurrentX
        curY = .CurrentY
        
        .CurrentX = "6.65in"
        xCol(7) = .CurrentX
        curY = .CurrentY
        
        .CurrentX = "7.4in"
        xCol(8) = .CurrentX
        curY = .CurrentY
        
        
        'Now spit out the headings at the indicated spacing
        Set cAxles = LoadingSummary(cTruck, bIsOnRoad)
        For i% = 1 To cAxles.Count
            .CurrentX = xCol(i% + 1)
            .Text = cAxles(i%).sDescription
        Next i%
        
        'Line =================================================
        .CurrentY = .CurrentY + 1.25 * vLineHt
        .DrawLine "1in", .CurrentY, .PageWidth - .MarginRight, .CurrentY
        .CurrentY = curY + 0.5 * vLineHt
        
        
        'Bridge Law Limit ===================================================
        If bIsOnRoad Then
            .CurrentY = .CurrentY + 1# * vLineHt
            .CurrentX = xCol(1)
            .FontBold = True
            .Text = "Bridge Law Limit"
            .FontBold = False
            'Only show Bridge Law limit for entire truck
            i% = cAxles.Count
            .CurrentX = xCol(i% + 1)
            sMsg = BridgeLawWarnings(cTruck.BridgeLaw, LoadedAxles(cTruck), dblTemp)
            dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblTemp, "#,##0")
        End If
        
        'Mfg Loading Limits ==============================================
        .CurrentY = .CurrentY + 1# * vLineHt
        .CurrentX = xCol(1)
        .FontBold = True
        .Text = "Mfg Limits"
        .FontBold = False
        For i% = 1 To cAxles.Count - 1
            .CurrentX = xCol(i% + 1)
            dblTemp = cAxles(i%).AllowableLd
            dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblTemp, "#,##0")
        Next i%
        'Make sure to show Mfg Limit for chassis
        .CurrentX = xCol(cAxles.Count + 1)
        If cTruck.Chassis.WtLimitTotal <> 0 Then
            dblTemp = cTruck.Chassis.WtLimitTotal
        Else
            dblTemp = cTruck.Chassis.WtLimitFront + cTruck.Chassis.WtLimitRear
        End If
        dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#,##0")
        
        'Actual Loading =================================================
        .CurrentY = .CurrentY + 1# * vLineHt
        .CurrentX = xCol(1)
        .FontBold = True
        .Text = "Calculated Loading"
        .FontBold = False
        For i% = 1 To cAxles.Count
            .CurrentX = xCol(i% + 1)
            dblTemp = cAxles(i%).ActualLd
            dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblTemp, "#,##0")
        Next i%
        
        'Over/Under Loading warnings =========================================
        'Only show warnings area if there are warnings
        sMsg = LoadingViolations(cTruck, bIsOnRoad)
        If sMsg <> "" Then
            .CurrentY = .CurrentY + 1.5 * vLineHt
            .CurrentX = xCol(1)
            .FontBold = True
            .Text = "Warnings for this Scenario "
            .FontBold = False
            If bIsOnRoad Then
                .Text = "(On Road): "
            Else
                .Text = "(Off Road): "
            End If
            .FontItalic = True
            sWarnings = Split(sMsg, vbCrLf)
            For i% = 0 To UBound(sWarnings)
                .CurrentY = .CurrentY + 1# * vLineHt
                .CurrentX = "2in"
                .Text = sWarnings(i%)
            Next i%
        End If
        
        'Logo in corner ===================================================
        .DrawPicture imgLogo, .MarginLeft, .PageHeight - .MarginBottom, "1in", ".24546in"
        
        'Tag Pressure ===================================================
        .FontSize = 8 'use smaller text
        .CalcParagraph = "H"
        vLineHt = .TextHei
        .CurrentY = "2.7in"
        .CurrentX = "5.5in"
        For i% = 1 To cAxles.Count
            If cAxles(i%).sDescription = "Pusher" Or cAxles(i%).sDescription = "Tag" _
             And cAxles(i%).ActualLd > 0 Then
                'This is a Tag axle of some sort
                .FontBold = True
                .Text = cAxles(i%).sDescription
                .FontBold = False
                .Text = " => "
                dblTemp = cAxles(i%).ActualLd
                dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
                .Text = Format(dblTemp, "#,##0 ") & cGlobalInfo.MassUnits.Display
                If cAxles(i%).ForceToPressure <> 0 Then
                    'Indicate Tag Air-Pressure if applicable
                    dblTemp = cAxles(i%).ActualLd * cAxles(i%).ForceToPressure
                    .Text = " [Set Air Pressure = " & Format(dblTemp, "#,##0") & "]"
                End If
                'Prep for next line
                .CurrentY = .CurrentY + vLineHt
                .CurrentX = "5.5in"
            End If
        Next i%
        
        .EndDoc
    End With
    
    Set cAxles = Nothing
End Sub


Public Sub ShopReport(cTruck As clsTruck)
    'This sub is called to create a report for Fred and the shop people.
    'It lists chassis dimensions and mounting location info.
    Dim i%
    Dim dblTemp As Double
    Dim sMsg As String
    Dim vLineHt
    Dim xCol(4) 'X-Position of column(i)
    Dim sDistFormat As String
    Dim dblTotalWt As Double
    
    mFile.Visible = False
    mOtherSide.Visible = False
    
    Me.Show
    
    vsDraw.Visible = False 'Don't show a drawing

    If cGlobalInfo.DistanceUnits.Display = "in" Then
        sDistFormat = "#0 "
    Else
        sDistFormat = "#0.## "
    End If

    With vsPrinter
        .StartDoc
        
        .MarginLeft = "0.5in"
        .MarginRight = "0.5in"
        .MarginTop = "0.5in"
        .MarginBottom = "0.5in"
        
        .FontName = "Tahoma"
        .FontSize = 12
        .CalcParagraph = "H"
        vLineHt = .TextHei
        
        .FontSize = 28
        .CurrentY = ".5in"
        .TextAlign = taCenterTop
        .Paragraph = cTruck.Description
        
        .FontSize = 12
        .Paragraph = "SN" & cTruck.SN
        .CurrentY = .CurrentY + vLineHt * 3
        
        'Chassis Information Header =================================================
        .FontSize = 21
        .TextAlign = taLeftTop
        .Paragraph = "Chassis Information"
        .CurrentY = .CurrentY + vLineHt / 2
        
        'Chassis Information Header =================================================
        .FontSize = 12
        xCol(1) = "1in"
        xCol(2) = "4.75in"
        
        .CurrentX = xCol(1)
        .FontBold = True
        .Text = "Description: "
        .FontBold = False
        .Text = cTruck.Chassis.FullName
        
        .CurrentY = .CurrentY + vLineHt * 1.25
        
        .CurrentX = xCol(1)
        .FontBold = True
        .Text = "Wheel Base"
        If cTruck.Chassis.TwinSteerSeparation > 0 Then
            .Text = "*: "
        Else
            .Text = ": "
        End If
        .FontBold = False
        dblTemp = (cTruck.Chassis.WB + cTruck.Chassis.TwinSteerSeparation / 2) * cGlobalInfo.DistanceUnits.Multiplier
        .Text = Format(dblTemp, sDistFormat) & " " & cGlobalInfo.DistanceUnits.Display
        
        .CurrentX = xCol(2)
        .FontBold = True
        .Text = "Front Axle"
        If cTruck.Chassis.TwinSteerSeparation > 0 Then
            .Text = "* to back of Cab: "
        Else
            .Text = " to back of Cab: "
        End If
        .FontBold = False
        dblTemp = cTruck.Chassis.BackOfCab * cGlobalInfo.DistanceUnits.Multiplier
        .Text = Format(dblTemp, sDistFormat) & " " & cGlobalInfo.DistanceUnits.Display
        
        .CurrentY = .CurrentY + vLineHt * 1.25
        
        .CurrentX = xCol(1)
        .FontBold = True
        .Text = "Empty Wt Front: "
        .FontBold = False
        dblTemp = cTruck.Chassis.WtFront * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
        
        .CurrentX = xCol(2)
        .FontBold = True
        .Text = "Mfg Wt Limit Front: "
        .FontBold = False
        dblTemp = cTruck.Chassis.WtLimitFront * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
        
        .CurrentY = .CurrentY + vLineHt * 1.25
        
        .CurrentX = xCol(1)
        .FontBold = True
        .Text = "Empty Wt Rear: "
        .FontBold = False
        dblTemp = cTruck.Chassis.WtRear * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
        
        .CurrentX = xCol(2)
        .FontBold = True
        .Text = "Mfg Wt Limit Rear: "
        .FontBold = False
        dblTemp = cTruck.Chassis.WtLimitRear * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
        
        .CurrentY = .CurrentY + vLineHt * 1.25
        
        .CurrentX = xCol(1)
        .FontBold = True
        .Text = "Tandem Spacing: "
        .FontBold = False
        dblTemp = cTruck.Chassis.TandemSpacing * cGlobalInfo.DistanceUnits.Multiplier
        .Text = Format(dblTemp, sDistFormat) & " " & cGlobalInfo.DistanceUnits.Display
        
        .CurrentX = xCol(2)
        .FontBold = True
        .Text = "Mfg Wt Limit Total: "
        .FontBold = False
        dblTemp = cTruck.Chassis.WtLimitTotal * cGlobalInfo.MassUnits.Multiplier
        .Text = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
        
        
        .CurrentY = .CurrentY + vLineHt * 2.5
        
        'Tag Axle Information =================================================
        .FontSize = 10
        .FontBold = False
        .CurrentX = "1in"
        xCol(1) = .CurrentX
        .Text = ""
        
        .CurrentX = "2in"
        xCol(2) = .CurrentX
        .Text = "(fm front Axle)"
        
        .FontSize = 12
        .FontBold = True
        .CurrentX = "3.1in"
        xCol(3) = .CurrentX
        .Text = "Tag Assy"
        
        .CurrentX = "4.15in"
        xCol(4) = .CurrentX
        .Text = "Manufacturers"
        
        .CurrentY = .CurrentY + vLineHt
        .CurrentX = xCol(1)
        .Text = "Tag Axle"
        
        .CurrentX = xCol(2)
        .Text = "Location"
        
        .CurrentX = xCol(3)
        .Text = "Weight"
        
        .CurrentX = xCol(4)
        .Text = "Loading Limit"
        
        'Line =================================================
        .CurrentY = .CurrentY + 1.25 * vLineHt
        .DrawLine "1in", .CurrentY, "5.625in", .CurrentY
                
        'Tags =================================================
        .FontBold = False
        .CurrentY = .CurrentY + 0.5 * vLineHt
        For i% = 1 To cTruck.Chassis.Tags.Count
            'FullName
            .CurrentX = xCol(1)
            If cTruck.Chassis.Tags(i%).Location < cTruck.Chassis.WB Then
                .Text = "Pusher"
            Else
                .Text = "Tag"
            End If
            'Location
            .CurrentX = xCol(2)
            dblTemp = (cTruck.Chassis.Tags(i%).Location + cTruck.Chassis.TwinSteerSeparation / 2) * cGlobalInfo.DistanceUnits.Multiplier
            .Text = Format(dblTemp, sDistFormat) & " " & cGlobalInfo.DistanceUnits.Display
            'Tag Wt
            .CurrentX = xCol(3)
            dblTemp = cTruck.Chassis.Tags(i%).Weight * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
            'Mfg Limit
            .CurrentX = xCol(4)
            dblTemp = cTruck.Chassis.Tags(i%).WtLimit * cGlobalInfo.MassUnits.Multiplier
            .Text = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
            
            .CurrentY = .CurrentY + 1# * vLineHt
        Next i%
        
        
        'Dual Steer Axle Notice ============================================
        If cTruck.Chassis.TwinSteerSeparation > 0 Then
            .CurrentY = .CurrentY + 1# * vLineHt
        
            .CurrentX = xCol(1)
            .FontBold = True
            .FontUnderline = True
            .Text = "*NOTE"
            
            .CurrentY = .CurrentY + 1# * vLineHt
            .CurrentX = xCol(1)
            .FontBold = False
            .FontUnderline = False
            .FontItalic = True
            .MarginLeft = xCol(1)
            sMsg = "This chassis has dual steer axles.  The indicated Wheelbase is measured " & _
                    "from the center of the front steering axle to the center of the rear " & _
                    "tandem.  Likewise, all other measurements relative to the 'Front Axle' " & _
                    " are referenced to the first steering axle."
            .Paragraph = sMsg
            .MarginLeft = "0.5in"
            .FontItalic = False
        End If
        
        
        'Body Information Header ============================================
        .CurrentY = .CurrentY + 2# * vLineHt
        .FontSize = 21
        .CurrentX = .MarginLeft
        .Paragraph = "Unibody Components Information"
        .CurrentY = .CurrentY + vLineHt / 2
        
        
        'List of Components ============================================
        'Start with Body
        .FontSize = 12
        .FontItalic = False
        .CurrentX = xCol(1)
        dblTemp = (cTruck.BodyLocation + cTruck.Chassis.TwinSteerSeparation / 2) * cGlobalInfo.DistanceUnits.Multiplier
        sMsg = Format(dblTemp, sDistFormat) & cGlobalInfo.DistanceUnits.Display
        sMsg = "Body [Located " & sMsg & " from front axle]"
        .Text = sMsg
        .FontSize = 10
        .FontItalic = True
        dblTemp = cTruck.Body.EmptyWeight * cGlobalInfo.MassUnits.Multiplier
        dblTotalWt = dblTemp
        sMsg = " -- (Empty Wt = " & Format(dblTemp, "#,##0 ") _
                & cGlobalInfo.MassUnits.Display & ")"
        .Text = sMsg
        .CurrentY = .CurrentY + 1 * vLineHt
        'Then loop thru components
        For i% = 1 To cTruck.Components.Count
                .CurrentX = xCol(1)
                .FontSize = 12
                .FontItalic = False
                .Text = cTruck.Components(i%).DisplayName
                .FontSize = 10
                .FontItalic = True
                dblTemp = cTruck.Components(i%).EmptyWeight * cGlobalInfo.MassUnits.Multiplier
                dblTotalWt = dblTotalWt + dblTemp
                sMsg = " -- (Empty Wt = " & Format(dblTemp, "#,##0 ") _
                        & cGlobalInfo.MassUnits.Display & ")"
                .Text = sMsg
                .CurrentY = .CurrentY + 1 * vLineHt
        Next i%
        
        'Now show the Chassis total Wt
        .CurrentY = .CurrentY + 1 * vLineHt
        sMsg = "Chassis "
        If cTruck.Chassis.Tags.Count > 0 Then sMsg = sMsg & "(and Tags) "
        .CurrentX = xCol(1)
        .FontSize = 12
        .FontItalic = False
        .Text = sMsg
        dblTemp = cTruck.Chassis.WtFront + cTruck.Chassis.WtRear
        For i% = 1 To cTruck.Chassis.Tags.Count
            dblTemp = dblTemp + cTruck.Chassis.Tags(i%).Weight
        Next i%
        dblTemp = dblTemp * cGlobalInfo.MassUnits.Multiplier
        dblTotalWt = dblTotalWt + dblTemp
        sMsg = " -- (Wt = " & Format(dblTemp, "#,##0 ") _
                & cGlobalInfo.MassUnits.Display & ")"
        .FontSize = 10
        .FontItalic = True
        .Text = sMsg
        .CurrentY = .CurrentY + 1 * vLineHt
        .CurrentY = .CurrentY + 1 * vLineHt
        
        'Now show the TOTAL empty Wt
        .FontSize = 12
        .FontItalic = False
        .CurrentX = xCol(1)
        .Text = "TOTAL Weight (incl. empty body, empty components, and chassis) = "
        .FontSize = 12
        .FontBold = True
        .Text = Format(dblTotalWt, "#,##0 ") & cGlobalInfo.MassUnits.Display
        .EndDoc
    End With
    
End Sub


Private Sub ShowDXFReport(cTruck As clsTruck, strFile As String, AspectRatio As Double)
    'This sub is called to create a loading report.  It shows a simplified picture
    ' of the truck along with loading information.
    Dim vLineHt
    Dim sWidth As String
    Dim sHeight As String
    
    mFile.Visible = True
    mOtherSide.Visible = True
    
    With vsDraw
        .PageWidth = 10 * 1440 '"10in"
        .PageHeight = .PageWidth / AspectRatio
        .Zoom = 0
        '.Picture = LoadPicture(strFile)
        .LoadDoc strFile
        .Refresh
    End With

    With vsPrinter
        .Clear
        .Orientation = orLandscape
        .StartDoc
        
        .MarginLeft = "0.5in"
        .MarginRight = "0.5in"
        .MarginTop = "0.5in"
        .MarginBottom = "0.5in"
        
        .FontName = "Tahoma"
        .FontSize = 12
        .CalcParagraph = "H"
        vLineHt = .TextHei
        
        .FontSize = 28
        .CurrentY = ".5in"
        .TextAlign = taCenterTop
        .Paragraph = cTruck.Description
        
        .FontSize = 12
        .Paragraph = "Unibody General Arrangement"
        .BrushColor = vbWhite
        sWidth = "9in"
        sHeight = CStr(9# / AspectRatio) & "in"
        .DrawPicture vsDraw.Picture, ".5in", .CurrentY, sWidth, sHeight

        .CurrentY = .Y2
        
        'Logo in corner ===================================================
        .DrawPicture imgLogo, .MarginLeft, .PageHeight - .MarginBottom, "1in", ".24546in"
        
        .EndDoc
    End With
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    'Fit printer object into window
    vsPrinter.Move vsPrinter.Left, vsPrinter.Top, ScaleWidth - 2 * vsPrinter.Left, ScaleHeight - vsPrinter.Left - vsPrinter.Top
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cCurTruck = Nothing
    DXFReader1.Clear
    
End Sub


Private Sub mSaveAs_Click()
    'Save the modified truck to a new file
    Dim sMsg As String
    Dim strNewFile As String
    Dim sFile As String
    Dim file_DXF As String
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
    
    file_DXF = AddBackslash(App.Path) & "~temp.dxf"

    FileCopy file_DXF, strNewFile
    Exit Sub
errHandler:
    ErrorIn "frmDXF.mSaveAs_Click"
End Sub

Public Sub ShowDrawing(cTruck As clsTruck)
    Set cCurTruck = cTruck
    bCurbSide = False
    Me.Show
    
    mOtherSide_Click
End Sub


Private Sub mOtherSide_Click()
    Dim file_DXF As String
    Dim file_WMF As String
    
    bCurbSide = Not bCurbSide
    'Change Menu text so user knows he must wait
    mOtherSide.Caption = "Wait while drawing is Rendering..."
    mOtherSide.Enabled = False
    mFile.Visible = False
    
    RenderTruck cCurTruck, DXFReader1, bCurbSide

    file_WMF = AddBackslash(App.Path) & "~temp.wmf"
    file_DXF = AddBackslash(App.Path) & "~temp.dxf"
    
    'fill a second control so that we can save the WMF
    AspectRatio = (DXFReader1.MaxX - DXFReader1.MinX) / (DXFReader1.MaxY - DXFReader1.MinY)
    
'    With DXFReader1
'        .Height = 12000
'        .Width = .Height * AspectRatio
'        .ZoomExtents
'        .SaveBMP file_WMF
'    End With

    With DXFReader2
        .Height = 12000
        .Width = .Height * AspectRatio
    End With
    mOtherSide.Caption = "Wait while DXF is created..."
    With DXFReader1
        .SetLimits .MinX, .MinY, .MaxX, .MaxY
        .WriteDXF file_DXF, True, 6, 10
    End With
    
    With DXFReader2
        .FileName = file_DXF
        .ZoomExtents
        mOtherSide.Caption = "Wait while WMF is created..."
        .SaveWMF file_WMF
   End With
 
    'Now Call the show an updated report
    ShowDXFReport cCurTruck, file_WMF, AspectRatio

    'Change Menu text so user knows he can now do something
    mFile.Visible = True
    If bCurbSide Then
        mOtherSide.Caption = "Curb Side"
    Else
        mOtherSide.Caption = "Street Side"
    End If
    mOtherSide.Enabled = True
End Sub

