Attribute VB_Name = "modDrawing"
Option Explicit

Private Const BORDERWIDTH As Single = 0.3 '(m) truck scale
Private Const TIREDIA As Single = 0.686 '(m) tire diameter [tag = 75%]
Private Const HUBRADIUS As Single = 150 '(mm) ~ 12in Dia
Private Const FRAMEHT As Single = 0.254 '(m) height of frame rail
Private Const FRM2AXLE As Single = 0.483 '(m) distance from frame rail top down to axle
Private Const REAREXTENSION As Single = 1.22 '(m) how far frame extends beyond last axle

Private NumTanks As Integer
Private CurPctEmul As Double
Private TankWidth() As Double 'ratio of body width
Private TankHeight As Double
Private sEmptyTank() As String 'Outlines of tanks (vector coords)
Private dblBodyLocation As Double '(m)
Private bStdMt As Boolean
Private colButtonCenters As Collection 'updated at


'ocDefaultColor is the same as "(Automatic)" in the property pages
Public Const ocDefaultColor As Long = &HFFFF

Public Const ocColorAliceBlue As Long = &HFFF8F0
Public Const ocColorAntiqueWhite As Long = &HD7EBFA
Public Const ocColorAquamarine As Long = &HD4FF7F
Public Const ocColorAzure As Long = &HFFFFF0
Public Const ocColorBeige As Long = &HDCF5F5
Public Const ocColorBisque As Long = &HC4E4FF
Public Const ocColorBlack As Long = &H0
Public Const ocColorBlanchedAlmond As Long = &HCDEBFF
Public Const ocColorBlue As Long = &HFF0000
Public Const ocColorBlueViolet As Long = &HE22B8A
Public Const ocColorBrown As Long = &H2A2AA5
Public Const ocColorBurlywood As Long = &H87B8DE
Public Const ocColorCadetBlue As Long = &HA09E5F

'ocColorChartreuse As Long = &H00FF7F
'The above would be true, but Visual Basic removes the leading zeros
Public Const ocColorChartreuse As Long = 65407

Public Const ocColorChocolate As Long = &H1E69D2
Public Const ocColorCoral As Long = &H507FFF
Public Const ocColorCornflowerBlue As Long = &HED9564
Public Const ocColorCornsilk As Long = &HDCF8FF
Public Const ocColorCyan As Long = &HFFFF00
Public Const ocColorDarkGoldenrod As Long = &HB86B8

'ocColorDarkGreen As Long = &H006400
'The above would be true, but Visual Basic removes the leading zeros
Public Const ocColorDarkGreen As Long = 25600

Public Const ocColorDarkKhaki As Long = &H6BB7BD
Public Const ocColorDarkOliveGreen As Long = &H2F6B55
Public Const ocColorDarkOrange As Long = &H8CFF
Public Const ocColorDarkOrchid As Long = &HCC3299
Public Const ocColorDarkSalmon As Long = &H7A96E9
Public Const ocColorDarkSeaGreen As Long = &H8FBC8F
Public Const ocColorDarkSlateBlue As Long = &H8B3D48
Public Const ocColorDarkSlateGray As Long = &H4F4F2F
Public Const ocColorDarkTurquoise As Long = &HD1CE00
Public Const ocColorDarkViolet As Long = &HD30094
Public Const ocColorDeepPink As Long = &H9314FF
Public Const ocColorDeepSkyBlue As Long = &HFFBF00
Public Const ocColorDodgerBlue As Long = &HFF901E
Public Const ocColorFirebrick As Long = &H2222B2
Public Const ocColorFloralWhite As Long = &HF0FAFF
Public Const ocColorForestGreen As Long = &H228B22
Public Const ocColorGainsboro As Long = &HDCDCDC
Public Const ocColorGhostWhite As Long = &HFFF8F8

Public Const ocColorGold As Long = 55295
Public Const ocColorGoldenrod As Long = &H20A5DA
Public Const ocColorGray As Long = &HBEBEBE
Public Const ocColorGray0 As Long = &H0
Public Const ocColorGray1 As Long = &H30303
Public Const ocColorGray2 As Long = &H50505
Public Const ocColorGray3 As Long = &H80808
Public Const ocColorGray4 As Long = &HA0A0A
Public Const ocColorGray5 As Long = &HD0D0D
Public Const ocColorGray6 As Long = &HF0F0F
Public Const ocColorGray7 As Long = &H121212
Public Const ocColorGray8 As Long = &H141414
Public Const ocColorGray9 As Long = &H171717
Public Const ocColorGray10 As Long = &H1A1A1A
Public Const ocColorGray11 As Long = &H1C1C1C
Public Const ocColorGray12 As Long = &H1F1F1F
Public Const ocColorGray13 As Long = &H212121
Public Const ocColorGray14 As Long = &H242424
Public Const ocColorGray15 As Long = &H262626
Public Const ocColorGray16 As Long = &H292929
Public Const ocColorGray17 As Long = &H2B2B2B
Public Const ocColorGray18 As Long = &H2E2E2E
Public Const ocColorGray19 As Long = &H303030
Public Const ocColorGray20 As Long = &H333333
Public Const ocColorGray21 As Long = &H363636
Public Const ocColorGray22 As Long = &H383838
Public Const ocColorGray23 As Long = &H3B3B3B
Public Const ocColorGray24 As Long = &H3D3D3D
Public Const ocColorGray25 As Long = &H404040
Public Const ocColorGray26 As Long = &H424242
Public Const ocColorGray27 As Long = &H454545
Public Const ocColorGray28 As Long = &H474747
Public Const ocColorGray29 As Long = &H4A4A4A
Public Const ocColorGray30 As Long = &H4D4D4D
Public Const ocColorGray31 As Long = &H4F4F4F
Public Const ocColorGray32 As Long = &H525252
Public Const ocColorGray33 As Long = &H545454
Public Const ocColorGray34 As Long = &H575757
Public Const ocColorGray35 As Long = &H595959
Public Const ocColorGray36 As Long = &H5C5C5C
Public Const ocColorGray37 As Long = &H5E5E5E
Public Const ocColorGray38 As Long = &H616161
Public Const ocColorGray39 As Long = &H636363
Public Const ocColorGray40 As Long = &H666666
Public Const ocColorGray41 As Long = &H696969
Public Const ocColorGray42 As Long = &H6B6B6B
Public Const ocColorGray43 As Long = &H6E6E6E
Public Const ocColorGray44 As Long = &H707070
Public Const ocColorGray45 As Long = &H737373
Public Const ocColorGray46 As Long = &H757575
Public Const ocColorGray47 As Long = &H787878
Public Const ocColorGray48 As Long = &H7A7A7A
Public Const ocColorGray49 As Long = &H7D7D7D
Public Const ocColorGray50 As Long = &H7F7F7F
Public Const ocColorGray51 As Long = &H828282
Public Const ocColorGray52 As Long = &H858585
Public Const ocColorGray53 As Long = &H878787
Public Const ocColorGray54 As Long = &H8A8A8A
Public Const ocColorGray55 As Long = &H8C8C8C
Public Const ocColorGray56 As Long = &H8F8F8F
Public Const ocColorGray57 As Long = &H919191
Public Const ocColorGray58 As Long = &H949494
Public Const ocColorGray59 As Long = &H969696
Public Const ocColorGray60 As Long = &H999999
Public Const ocColorGray61 As Long = &H9C9C9C
Public Const ocColorGray62 As Long = &H9E9E9E
Public Const ocColorGray63 As Long = &HA1A1A1
Public Const ocColorGray64 As Long = &HA3A3A3
Public Const ocColorGray65 As Long = &HA6A6A6
Public Const ocColorGray66 As Long = &HA8A8A8
Public Const ocColorGray67 As Long = &HABABAB
Public Const ocColorGray68 As Long = &HADADAD
Public Const ocColorGray69 As Long = &HB0B0B0
Public Const ocColorGray70 As Long = &HB3B3B3
Public Const ocColorGray71 As Long = &HB5B5B5
Public Const ocColorGray72 As Long = &HB8B8B8
Public Const ocColorGray73 As Long = &HBABABA
Public Const ocColorGray74 As Long = &HBDBDBD
Public Const ocColorGray75 As Long = &HBFBFBF
Public Const ocColorGray76 As Long = &HC2C2C2
Public Const ocColorGray77 As Long = &HC4C4C4
Public Const ocColorGray78 As Long = &HC7C7C7
Public Const ocColorGray79 As Long = &HC9C9C9
Public Const ocColorGray80 As Long = &HCCCCCC
Public Const ocColorGray81 As Long = &HCFCFCF
Public Const ocColorGray82 As Long = &HD1D1D1
Public Const ocColorGray83 As Long = &HD4D4D4
Public Const ocColorGray84 As Long = &HD6D6D6
Public Const ocColorGray85 As Long = &HD9D9D9
Public Const ocColorGray86 As Long = &HDBDBDB
Public Const ocColorGray87 As Long = &HDEDEDE
Public Const ocColorGray88 As Long = &HE0E0E0
Public Const ocColorGray89 As Long = &HE3E3E3
Public Const ocColorGray90 As Long = &HE5E5E5
Public Const ocColorGray91 As Long = &HE8E8E8
Public Const ocColorGray92 As Long = &HEBEBEB
Public Const ocColorGray93 As Long = &HEDEDED
Public Const ocColorGray94 As Long = &HF0F0F0
Public Const ocColorGray95 As Long = &HF2F2F2
Public Const ocColorGray96 As Long = &HF5F5F5
Public Const ocColorGray97 As Long = &HF7F7F7
Public Const ocColorGray98 As Long = &HFAFAFA
Public Const ocColorGray99 As Long = &HFCFCFC

Public Const ocColorGreen As Long = 65280
Public Const ocColorGreenYellow As Long = &H2FFFAD
Public Const ocColorHoneydew As Long = &HF0FFF0
Public Const ocColorHotPink As Long = &HB469FF
Public Const ocColorIndianRed As Long = &H5C5CCD
Public Const ocColorIvory As Long = &HF0FFFF
Public Const ocColorKhaki As Long = &H8CE6F0
Public Const ocColorLavender As Long = &HFAE6E6
Public Const ocColorLavenderBlush As Long = &HF5F0FF

Public Const ocColorLawnGreen As Long = 64636
Public Const ocColorLemonChiffon As Long = &HCDFAFF
Public Const ocColorLightBlue As Long = &HE6D8AD
Public Const ocColorLightCoral As Long = &H8080F0
Public Const ocColorLightCyan As Long = &HFFFFE0
Public Const ocColorLightGoldenrod As Long = &H82DDEE
Public Const ocColorLightGoldenrodYellow As Long = &HD2FAFA
Public Const ocColorLightGray As Long = &HD3D3D3
Public Const ocColorLightPink As Long = &HC1B6FF
Public Const ocColorLightSalmon As Long = &H7AA0FF
Public Const ocColorLightSeaGreen As Long = &HAAB220
Public Const ocColorLightSkyBlue As Long = &HFACE87
Public Const ocColorLightSlateBlue As Long = &HFF7084
Public Const ocColorLightSlateGray As Long = &H998877
Public Const ocColorLightSteelBlue As Long = &HDEC4B0
Public Const ocColorLightYellow As Long = &HE0FFFF
Public Const ocColorLimeGreen As Long = &H32CD32
Public Const ocColorLinen As Long = &HE6F0FA
Public Const ocColorMagenta As Long = &HFF00FF
Public Const ocColorMaroon As Long = &H6030B0
Public Const ocColorMediumAquamarine As Long = &HAACD66
Public Const ocColorMediumBlue As Long = &HCD0000
Public Const ocColorMediumOrchid As Long = &HD355BA
Public Const ocColorMediumPurple As Long = &HDB7093
Public Const ocColorMediumSeaGreen As Long = &H71B33C
Public Const ocColorMediumSlateBlue As Long = &HEE687B
Public Const ocColorMediumSpringGreen As Long = &H9AFA00
Public Const ocColorMediumTurquoise As Long = &HCCD148
Public Const ocColorMediumVioletRed As Long = &H8515C7
Public Const ocColorMidnightBlue As Long = &H701919
Public Const ocColorMintCream As Long = &HFAFFF5
Public Const ocColorMistyRose As Long = &HE1E4FF
Public Const ocColorMoccasin As Long = &HB5E4FF
Public Const ocColorNavajoWhite As Long = &HADDEFF
Public Const ocColorNavyBlue As Long = &H800000
Public Const ocColorOldLace As Long = &HE6F5FD
Public Const ocColorOliveDrab As Long = &H238E6B

Public Const ocColorOrange As Long = 42495
Public Const ocColorOrangeRed As Long = 17919

Public Const ocColorOrchid As Long = &HD670DA
Public Const ocColorPaleGoldenrod As Long = &HAAE8EE
Public Const ocColorPaleGreen As Long = &H98FB98
Public Const ocColorPaleTurquoise As Long = &HEEEEAF
Public Const ocColorPaleVioletRed As Long = &H9370DB
Public Const ocColorPapayaWhip As Long = &HD5EFFF
Public Const ocColorPeachPuff As Long = &HB9DAFF
Public Const ocColorPeru As Long = &H3F85CD
Public Const ocColorPink As Long = &HCBC0FF
Public Const ocColorPlum As Long = &HDDA0DD
Public Const ocColorPowderBlue As Long = &HE6E0B0
Public Const ocColorPurple As Long = &HF020A0
Public Const ocColorRed As Long = &HFF
Public Const ocColorRosyBrown As Long = &H8F8FBC
Public Const ocColorRoyalBlue As Long = &HE16941
Public Const ocColorSaddleBrown As Long = &H13458B
Public Const ocColorSalmon As Long = &H7280FA
Public Const ocColorSandyBrown As Long = &H60A4F4
Public Const ocColorSeaGreen As Long = &H578B2E
Public Const ocColorSeashell As Long = &HEEF5FF
Public Const ocColorSienna As Long = &H2D52A0
Public Const ocColorSkyBlue As Long = &HEBCE87
Public Const ocColorSlateBlue As Long = &HCD5A6A
Public Const ocColorSlateGray As Long = &H908070
Public Const ocColorSnow As Long = &HFAFAFF
Public Const ocColorSpringGreen As Long = &H7FFF00
Public Const ocColorSteelBlue As Long = &HB48246
Public Const ocColorTan As Long = &H8CB4D2
Public Const ocColorThistle As Long = &HD8BFD8
Public Const ocColorTomato As Long = &H4763FF
Public Const ocColorTurquoise As Long = &HD0E040
Public Const ocColorViolet As Long = &HEE82EE
Public Const ocColorVioletRed As Long = &H9020D0
Public Const ocColorWheat As Long = &HB3DEF5
Public Const ocColorWhite As Long = &HFFFFFF

Public Const ocColorYellow As Long = 65535
Public Const ocColorYellowGreen As Long = &H32CD9A

Public Function UpdateDrawing(vsdBody As vsDraw, cTruck As clsTruck)
    'Call this sub to show/update tank contents
    Dim i%
    Dim dblPctFull As Double
    
    If NumTanks = 0 Or vsdBody.Tag <> cTruck.SN Then
        'Need to Initialize first
        InitializeDrawing vsdBody, cTruck
    End If

    'Show Tanks
    For i% = 1 To NumTanks
        'clear tank
        vsdBody.BrushColor = vsdBody.BackColor
        vsdBody.Polygon = OutlineString(vsdBody, i%, 1#)
        'Filled tank
        'dblPctFull = cTruck.Body.Tanks(i%).CurVol / cTruck.Body.Tanks(i%).Volume
        dblPctFull = (cTruck.Body.Tanks(i%).MaxHt - cTruck.Body.Tanks(i%).CurStkHt) / cTruck.Body.Tanks(i%).MaxHt
        
        If cTruck.Body.Tanks(i%).CurTankUse = ttEmulsion Then
            vsdBody.BrushColor = vbBlue 'Emulsion
        Else
            vsdBody.BrushColor = vbRed 'AN
        End If
        vsdBody.Polygon = OutlineString(vsdBody, i%, dblPctFull)
        'Tank Outline
        vsdBody.Polyline = sEmptyTank(i%)
    Next i%

    vsdBody.Refresh
    DoEvents
End Function


Private Function InitializeDrawing(vsdBody As vsDraw, cTruck As clsTruck)
    'Clear drawing, calculate basics, and draw chassis & empty body
    Dim i%
    Dim TotalVol As Double 'Total Volume capacity of Truck
    Dim dblWidth As Double 'basic drawing width
    Dim dblHeight As Double 'basic drawing width
    Dim a$
    Dim BodyWidth As Double
    Dim dblTemp As Double
    Dim dblTankCtr As Double
    Dim sFRL_X As String 'FrameRailLeft X
    Dim sFRR_X As String 'FrameRailRight X
    
    vsdBody.Clear 'clears drawing
    vsdBody.Tag = cTruck.SN 'stamp the drawing w/ truck SN
    
    NumTanks = cTruck.Body.Tanks.Count
    ReDim sEmptyTank(1 To NumTanks)
    ReDim TankWidth(1 To NumTanks)
        
    Set colButtonCenters = New Collection
    
    BodyWidth = cTruck.Body.BodyLength
    TankHeight = cTruck.Body.Tanks(1).MaxHt
    
    'Calculate simplistic estimated tankwidths by Vol % of Body width
    For i% = 1 To cTruck.Body.Tanks.Count
        TotalVol = TotalVol + cTruck.Body.Tanks(i%).Volume
    Next i%
    For i% = 1 To cTruck.Body.Tanks.Count
        TankWidth(i%) = BodyWidth * (cTruck.Body.Tanks(i%).Volume / TotalVol)
    Next i%
    
    'Calculate total width of diagram
    dblWidth = 0
    dblTemp = cTruck.Chassis.WB + cTruck.Chassis.TandemSpacing / 2
    If cTruck.Chassis.Tags.Count > 0 Then
        For i% = 1 To cTruck.Chassis.Tags.Count
            If cTruck.Chassis.Tags(i%).Location > dblTemp Then
                dblWidth = cTruck.Chassis.Tags(i%).Location
                'Add rear and forward extensions
                dblWidth = dblWidth + TIREDIA / 2 + REAREXTENSION
            End If
        Next i%
    End If
    If dblWidth = 0 Then
        'tags don't define extreme axles
        dblWidth = dblTemp + TIREDIA / 2 + REAREXTENSION
    End If
    dblTemp = cTruck.BodyLocation + BodyWidth + TIREDIA / 2
    If dblTemp > dblWidth Then
        'body extends beyond rear-most wheel, so use last compartment as the extreme right
        dblWidth = dblTemp
    End If
    'Allow for Twin Steer
    dblWidth = dblWidth + cTruck.Chassis.TwinSteerSeparation / 2
    
    'calculate total height of diagram
    dblHeight = TankHeight + FRM2AXLE + TIREDIA / 2
    
    With vsdBody
        .ScaleWidth = 1000# * (dblWidth + BORDERWIDTH * 2)
        .ScaleLeft = 1000# * (-BORDERWIDTH - TIREDIA / 2 - cTruck.Chassis.TwinSteerSeparation / 2)
        .ScaleHeight = -1000# * (dblHeight + BORDERWIDTH * 2)
        .ScaleTop = 1000# * (TankHeight + BORDERWIDTH)
        .Clear
        ' set brush
        .BrushStyle = bsSolid
    End With
    
    bStdMt = cTruck.IsStandardMount
    dblBodyLocation = cTruck.BodyLocation
    
    'draw framerail
    vsdBody.BrushColor = ocColorGray60
    sFRL_X = CStr(-cTruck.Chassis.TwinSteerSeparation / 2 * 1000)
    sFRR_X = CStr(1000 * (dblWidth - TIREDIA / 2 - cTruck.Chassis.TwinSteerSeparation / 2))
    a$ = sFRL_X & " 0 " & _
         sFRR_X & " 0 " & _
         sFRR_X & " " & CStr(1000 * -FRAMEHT) & " " & _
         sFRL_X & " " & CStr(1000 * -FRAMEHT) & " " & _
         sFRL_X & " 0"
    vsdBody.Polygon = a$
    'draw front tire(s)
    If cTruck.Chassis.TwinSteerSeparation = 0 Then
        'Single front Tire
        vsdBody.BrushColor = vbBlack
        vsdBody.DrawCircle 0, -1000 * FRM2AXLE, 1000 * TIREDIA / 2
        vsdBody.BrushColor = vbWhite
        vsdBody.DrawCircle 0, -1000 * FRM2AXLE, HUBRADIUS
    Else
        'Twin-steer
        vsdBody.BrushColor = vbBlack
        vsdBody.DrawCircle (cTruck.Chassis.TwinSteerSeparation / 2) * 1000, _
                            -1000 * FRM2AXLE, 1000 * TIREDIA / 2
        vsdBody.BrushColor = vbWhite
        vsdBody.DrawCircle (cTruck.Chassis.TwinSteerSeparation / 2) * 1000, _
                            -1000 * FRM2AXLE, HUBRADIUS
        vsdBody.BrushColor = vbBlack
        vsdBody.DrawCircle (-cTruck.Chassis.TwinSteerSeparation / 2) * 1000, _
                            -1000 * FRM2AXLE, 1000 * TIREDIA / 2
        vsdBody.BrushColor = vbWhite
        vsdBody.DrawCircle (-cTruck.Chassis.TwinSteerSeparation / 2) * 1000, _
                            -1000 * FRM2AXLE, HUBRADIUS
    End If
    
    'draw rear tire(s)
    If cTruck.Chassis.TandemSpacing > 0 Then
        'two wheels
        vsdBody.BrushColor = vbBlack
        vsdBody.DrawCircle 1000 * (cTruck.Chassis.WB - cTruck.Chassis.TandemSpacing / 2), -1000 * FRM2AXLE, 1000 * TIREDIA / 2
        vsdBody.DrawCircle 1000 * (cTruck.Chassis.WB + cTruck.Chassis.TandemSpacing / 2), -1000 * FRM2AXLE, 1000 * TIREDIA / 2
        vsdBody.BrushColor = vbWhite
        vsdBody.DrawCircle 1000 * (cTruck.Chassis.WB - cTruck.Chassis.TandemSpacing / 2), -1000 * FRM2AXLE, HUBRADIUS
        vsdBody.DrawCircle 1000 * (cTruck.Chassis.WB + cTruck.Chassis.TandemSpacing / 2), -1000 * FRM2AXLE, HUBRADIUS
    Else
        'one wheel
        vsdBody.DrawCircle 1000 * cTruck.Chassis.WB, -1000 * FRM2AXLE, 1000 * TIREDIA / 2
        vsdBody.BrushColor = vbWhite
        vsdBody.DrawCircle 1000 * cTruck.Chassis.WB, -1000 * FRM2AXLE, HUBRADIUS
    End If
    'draw tag(s)
    If cTruck.Chassis.Tags.Count > 0 Then
        For i% = 1 To cTruck.Chassis.Tags.Count
            
            If cTruck.Chassis.Tags(i%).DownwardForce = 0 Then
                dblTemp = -1000 * (FRM2AXLE - 0.125 * TIREDIA)
            Else
                dblTemp = -1000 * (FRM2AXLE + 0.125 * TIREDIA)
            End If
            vsdBody.BrushColor = vbBlack
            vsdBody.DrawCircle 1000 * cTruck.Chassis.Tags(i%).Location, dblTemp, 1000 * TIREDIA / 2 * 0.75
            vsdBody.BrushColor = vbWhite
            vsdBody.DrawCircle 1000 * cTruck.Chassis.Tags(i%).Location, dblTemp, HUBRADIUS * 0.75
        Next i%
    End If
    
    'Create Empty Tank Outlines
    For i% = 1 To NumTanks
        sEmptyTank(i%) = OutlineString(vsdBody, i%, 1#, dblTankCtr)
        vsdBody.Polyline = sEmptyTank(i%)
        colButtonCenters.Add dblTankCtr, CStr(i%)
    Next i%
End Function


Private Function OutlineString(vsdBody As vsDraw, TankNum As Integer, PctFull As Double, Optional TankCtr As Double) As String
    'Returns a string used to create PolyLine or PolyGon for the
    'indicated tank.  'dblBodyLocation' and 'bStdMt' are used
    'for accurate rendering of driver-side
    ' 'TankCtr' is returned for use as needed
    Dim i%
    Dim dblLeft As Double
    Dim dblRight As Double
    Dim dblTop As Double
    Dim dblBottom As Double
    Dim sOutline As String
    Dim ChassisHt As Double 'mm
    Dim dblBodyWidth As Double
    
    
    For i% = 1 To TankNum - 1
        dblLeft = dblLeft + TankWidth(i%)
        dblBodyWidth = dblBodyWidth + TankWidth(i%)
    Next i%
    
    
    If bStdMt Then
        dblLeft = dblBodyLocation + dblLeft
        dblRight = dblLeft + TankWidth(i%)
    Else
        dblLeft = dblBodyLocation + dblBodyWidth - dblLeft
        dblRight = dblLeft - TankWidth(i%)
    End If
    dblLeft = Round(1000# * dblLeft, 0)
    dblRight = Round(1000# * dblRight, 0)
    
    dblBottom = 0
    dblTop = Round(1000# * PctFull * TankHeight, 0)

    sOutline = CStr(dblLeft) & " " & CStr(dblBottom) & " , "
    sOutline = sOutline & CStr(dblRight) & " " & CStr(dblBottom) & " , "
    sOutline = sOutline & CStr(dblRight) & " " & CStr(dblTop) & " , "
    sOutline = sOutline & CStr(dblLeft) & " " & CStr(dblTop) & " , "
    sOutline = sOutline & CStr(dblLeft) & " " & CStr(dblBottom)
    
    TankCtr = ((dblLeft + dblRight) / 2 - vsdBody.ScaleLeft) / vsdBody.ScaleWidth
    
    OutlineString = sOutline
End Function

Public Function ButtonCenters() As Collection
    'Returns a collection of numbers that represent the centers of the tanks.
    ' The values are a percentage of the vsDraw width
    Set ButtonCenters = colButtonCenters
End Function
