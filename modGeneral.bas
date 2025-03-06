Attribute VB_Name = "modGeneral"
Option Explicit

Public Enum PlacementAllowableLocation
    paEitherSide = 0
    paStreetSideStd = 1
    paCurbSideStd = 2
    paCenter = 3
End Enum

Public Enum PlacementLocation
    plNotPlaced = 0
    plStreetSideStd = 1
    plCurbSideStd = 2
    plCenter = 3
End Enum

Public cActiveTruck As clsTruck

Public bInHouseVersion As Boolean
Public bHasBeenEdited As Boolean
Public bCanCreateALObjects As Boolean 'FALSE => User will not be able to create & save new AL Objects


'---------------------------------------------
'NOTES
'
'Version Compatability:
'   Revision-level changes: -All files of same Major.Minor level will work with any
'                            program of the same Major.Minor level.
'                           -Revision-level changes are basically for aesthetic changes
'                            to the GUI.
'
'   Minor-Level changes:    -The new program will read current and older files. Some
'                            of the new program features may simply not be usable.
'                            (e.g. new version may include auto-drawing, but old
'                            file won't have info for the drawings)
'                           -If user opens an older file, he will be told that the file
'                            can be read, but that he should upgrade the program for
'                            access to new features.
'                           -Files of same Major level but newer minor level will contain
'                            extra data that the (older) program will not be able to use.
'                            Users will not be able to save files unless they upgrade
'                            to the latest software (otherwise certain data in the file
'                            could be lost when file is saved(recreated) bt the program.
'
'   Major-Level changes:    -Should be able to read older files.  A warning will indicate
'                            that they should use the older program version with this file
'                            for guaranteed compatability.  Saving of older Major-level
'                            files will be prohibited.
'                           -Files of a newer major level will NOT be read.  The user will
'                            be told to get the newest version of software to read the
'                            file in question.
'---------------------------------------------


Public Sub Main()
    InitCommonControlsVB
    'Determine if this person can be creating & saving new AL Objects
    SaveSetting "AxleLoad_InHouse", "Options", "CanCreateALObjects", CBool(False)
    bCanCreateALObjects = CBool(GetSetting(appname:="AxleLoad_InHouse", _
                                Section:="Options", _
                                Key:="CanCreateALObjects", Default:="False"))
    
    If App.Title = "AxleLoad_InHouse" Then
        bInHouseVersion = True
    Else
        bInHouseVersion = False
    End If
    frmMain.Show
End Sub


Public Function IsNumber(InNum As Variant) As Boolean
    Dim NumOnly As Boolean
    NumOnly = False
    On Error GoTo NotNum2
    If InNum = val(InNum) Then
        'Value is a number (not an equation)
        NumOnly = True
    End If
NotNum2:
    On Error GoTo 0
    IsNumber = NumOnly
End Function

Public Function Var2Dbl(vIn As Variant) As Double
    'Like CDbl, but doesn't crash on empty variant
    On Error Resume Next
    
    Err.Clear
    Var2Dbl = CDbl(vIn)
    If Err.Number <> 0 Then
        Var2Dbl = 0#
    End If
End Function

Public Sub Log(strText As String)
    'This sub appends strText to a log text file
    Static sFl As String
    Dim intFlNum  As Integer
    
    If sFl = "" Then
        sFl = App.Path
        If Right$(sFl, 1) <> "\" Then
            sFl = sFl & "\"
        End If
        sFl = sFl & "LogFile.txt"
    End If
    
    intFlNum = FreeFile
    Open sFl For Append As intFlNum
    Print #intFlNum, strText
    Close intFlNum
End Sub

Public Function SummaryText(cTruck As clsTruck) As String
    '*** Used during Development.  Should be erased at time of release.
    Dim x%
    Dim sText As String
    Dim sProd As String
    Dim dblTemp As Double
    Dim strTemp As String

    'Show the good distribution
    sText = sText & "Valid Truck Configuration" & vbCrLf
    
     dblTemp = (cTruck.Body.MassANTotal + cTruck.Body.MassEmulTotal) * cGlobalInfo.MassUnits.Multiplier
     strTemp = Format(dblTemp, "###0 ") & cGlobalInfo.MassUnits.Display
    sText = sText & " Total Product = " & strTemp & vbCrLf
    
     dblTemp = 100# * cTruck.Body.MassEmulTotal / (cTruck.Body.MassANTotal / (1# - dblDFOPct) + cTruck.Body.MassEmulTotal)
     strTemp = Format(dblTemp, "#0") & "%"
    sText = sText & " %Emulsion = " & strTemp & vbCrLf
    
    For x% = 1 To cTruck.Body.Tanks.Count
        If cTruck.Body.Tanks(x%).CurTankUse = ttAN Then
            sProd = "AN"
        Else
            sProd = "Emulsion"
        End If
        dblTemp = cTruck.Body.Tanks(x%).CurVol * cTruck.Body.Tanks(x%).DensityContents * cGlobalInfo.MassUnits.Multiplier
        strTemp = Format(dblTemp, "#0 ") & cGlobalInfo.MassUnits.Display
        sText = sText & " Tank " & x% & " = (" & strTemp & _
               ") of " & sProd & vbTab & " (StkHt = " & Format(cTruck.Body.Tanks(x%).CurStkHt * cGlobalInfo.DistanceUnits.Multiplier, "#0 ") _
               & cGlobalInfo.DistanceUnits.Display & ")  Density =" & cTruck.Body.Tanks(x%).DensityContents & _
               "  %Full = " & Round(100 * cTruck.Body.Tanks(x%).CurVol / cTruck.Body.Tanks(x%).Volume, 0) & vbCrLf
    Next x%
    
    sText = sText & "Front Loading = " & Round(FrontLoading(cTruck) * cGlobalInfo.MassUnits.Multiplier, 0) & cGlobalInfo.MassUnits.Display & vbCrLf
    sText = sText & "Rear Loading = " & Round(RearLoading(cTruck) * cGlobalInfo.MassUnits.Multiplier, 0) & cGlobalInfo.MassUnits.Display & vbCrLf
    
    sText = sText & "-----------------------------------------" & vbCrLf & vbCrLf

    SummaryText = sText
End Function

Function AddBackslash(Path As String, Optional Char As String = "\") As String
'Append a backslash (or any character) at the end of a path if it isn't there already
    If Right$(Path, 1) <> Char Then
        AddBackslash = Path & Char
    Else
        AddBackslash = Path
    End If
End Function

Function GetFileBaseName(ByVal FileName As String, Optional ByVal IncludePath As Boolean) As String
'Retrieve a file's base name. If the second argument is true, the result include the file's path
    Dim i As Long, startPos As Long, endPos As Long
    
    startPos = 1
    
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case "."
                ' we've found the extension
                If IncludePath Then
                    ' if we must return the path, we've done
                    GetFileBaseName = Left$(FileName, i - 1)
                    Exit Function
                End If
                ' else, just take note of where the extension begins
                If endPos = 0 Then endPos = i - 1
            Case ":", "\"
                If Not IncludePath Then startPos = i + 1
                Exit For
        End Select
    Next
    
    If endPos = 0 Then
        ' this file has no extension
        GetFileBaseName = Mid$(FileName, startPos)
    Else
        GetFileBaseName = Mid$(FileName, startPos, endPos - startPos + 1)
    End If
End Function

Public Function PathlessFileName(FileName As String) As String
    Dim i As Long, startPos As Long, endPos As Long
    
    startPos = 1
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case ":", "\", "/"
                startPos = i + 1
                Exit For
        End Select
    Next
    PathlessFileName = Mid$(FileName, startPos)
End Function



Public Function FormattedLevel(cObject As Variant) As String
    'This function returns a string that contains the 'level' of the indicated
    ' component/tank.  If the object doesn't have a 'level' (e.g. a 'driver'), then
    ' an empty string is returned
    Dim dblTemp As Double
    Dim dblDensity As Double
    Dim sRtn As String
    On Error GoTo errHandler
    
    sRtn = ""
    
    If TypeName(cObject) = "clsComponent" Then
        If Not cObject.Capacity.UsesSightGauge Then
            'not sight-gauge, return Stick-Height or Weight
            If cObject.Capacity.StickLength.Count > 0 Then
                'Return Stick Height
                dblTemp = Round(cObject.Capacity.CurStkHt / 0.0254) * 0.0254 'to nearest inch
                dblTemp = dblTemp * cGlobalInfo.DistanceUnits.Multiplier
                sRtn = Format(dblTemp, "0.## ") & cGlobalInfo.DistanceUnits.Display
            Else
                'Return Weight
                dblTemp = cObject.Capacity.CurVol * cObject.Capacity.DensityContents _
                          * cGlobalInfo.MassUnits.Multiplier
                sRtn = Format(dblTemp, "# ") & cGlobalInfo.MassUnits.Display
            End If
        Else
            If cObject.Capacity.DensityContents <> 0 Then
                'must have some value for density if there are 'contents'
                dblTemp = cObject.Capacity.CurVol / cObject.Capacity.Volume
                'sRtn = Decimal2Fraction(dblTemp, 8) & " Full"
                sRtn = Format(dblTemp * 100, "#0") & "% Full"
            End If
        End If
    ElseIf TypeName(cObject) = "clsTank" Then
        'Return Stick Height
        dblTemp = Round(cObject.CurStkHt / 0.0254) * 0.0254 'to nearest inch
        dblTemp = dblTemp * cGlobalInfo.DistanceUnits.Multiplier
        sRtn = Format(dblTemp, "0.## ") & cGlobalInfo.DistanceUnits.Display
    Else
        Exit Function
    End If
    
    
    FormattedLevel = sRtn
    Exit Function
errHandler:
    ErrorIn "modGeneral.FormattedLevel(cObject)", cObject
End Function


Public Function Decimal2Fraction(ByVal val As Double, Optional Denominator As Double = -1) As String
   Dim df As Double
   Dim lUpperPart As Long
   Dim lLowerPart As Long
   
    If Denominator > 0 Then
        val = Round(val * Denominator, 0) / Denominator
    End If
   
    lUpperPart = 1
    lLowerPart = 1
   
    df = lUpperPart / lLowerPart
    While (df <> val)
        If (df < val) Then
            lUpperPart = lUpperPart + 1
        Else
            lLowerPart = lLowerPart + 1
            lUpperPart = val * lLowerPart
        End If
        df = lUpperPart / lLowerPart
    Wend
    Decimal2Fraction = CStr(lUpperPart) & "/" & CStr(lLowerPart)
End Function


Public Function FileExists(FileName As String) As Boolean
    'Return True if a file exists
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function


Public Function FixedFileName(ByRef sFile As String) As String
    'This function replaces invalid characters with spaces
    Dim sChar() As String
    Dim InvalidChars As String
    Dim i%
    
    InvalidChars = "\ / : * ? < > | " & Chr$(34)
    sChar = Split(InvalidChars, " ")
    
    For i% = 0 To UBound(sChar)
        sFile = Replace(sFile, sChar(i%), "")
    Next i%
    FixedFileName = sFile
End Function
