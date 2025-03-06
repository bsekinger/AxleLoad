Attribute VB_Name = "modAxleLoading"
Option Explicit

Public Function FrontLoading(cTruck As clsTruck) As Double
    'Sticklength, CG, and weight should have been calculated prior to
    'calling this function

    'Standard Mount
    '--------------
    '   FrontAxleLoad = sum[ Cwt(WB - l - Ccg)] / WB
    '       where Cwt = Component Weight
    '             Ccg = Dist from Component CG to Body Origin
    '               l = Dist from front axle to Front of Body
    '              WB = Dist between front and rear axle
    
    'Reverse Mount
    '-------------
    '   FrontAxleLoad = sum[ Cwt(WB - (l + BodyLength - Ccg))] / WB
    '       where Cwt = Component Weight
    '             Ccg = Dist from Component CG to Body Origin
    '               l = Dist from front axle to Front of Body
    '              WB = Dist between front and rear axle
    
    Dim dblSumMoments As Double
    Dim eM As Double 'Moment due to equipment weight
    Dim cM As Double 'Moment due to contents weight only
    Dim Cwt As Double
    Dim Ccg As Double
    Dim L As Double
    Dim WB As Double
    Dim BodyLength As Double
    Dim dblDensity As Double
    Dim cTank As clsTank
    Dim cComponent As clsComponent
    Dim i%
    Dim lmt As Integer

    BodyLength = cTruck.Body.BodyLength
    WB = cTruck.Chassis.WB
    L = cTruck.BodyLocation
    
    'Sum the Tank Loads -------------------
    'Contents
    lmt = cTruck.Body.Tanks.Count
    For i% = 1 To lmt
        Set cTank = cTruck.Body.Tanks(i%)
        dblDensity = cTank.DensityContents
        'calculate weight using density
        Ccg = cTank.CurContentCG 'calc'd elsewhere
        Cwt = cTank.CurVol * dblDensity
    
        If cTruck.IsStandardMount Then
            cM = Cwt * (WB - L - Ccg)
        Else
            cM = Cwt * (WB - (L + BodyLength - Ccg))
        End If
        'Sum Moments
        dblSumMoments = dblSumMoments + cM
    Next i%
    'Equipment
    Ccg = cTruck.Body.EmptyCG
    Cwt = cTruck.Body.EmptyWeight
    If cTruck.IsStandardMount Then
        eM = Cwt * (WB - L - Ccg)
    Else
        eM = Cwt * (WB - (L + BodyLength - Ccg))
    End If
    'Sum Moments
    dblSumMoments = dblSumMoments + eM
    
    'Sum the Component Loads --------------
    lmt = cTruck.Components.Count
    For i% = 1 To lmt
        Set cComponent = cTruck.Components(i%)
        'Equipment
        Cwt = cComponent.EmptyWeight
        Ccg = cComponent.EmptyCG + cComponent.Offset
        If cComponent.LocationReference = orBodyOrigin Then
            If cTruck.IsStandardMount Then
                eM = Cwt * (WB - L - Ccg)
            Else
                eM = Cwt * (WB - (L + BodyLength - Ccg))
            End If
        ElseIf cComponent.LocationReference = orFrontAxle Then
            eM = Cwt * (WB - Ccg)
        End If
        'Contents
        dblDensity = cComponent.Capacity.DensityContents
        If Var2Dbl(cComponent.Capacity.DefaultVolContents) > 0# Then
            'calculate using Default(user) content Volume
            Ccg = cComponent.Capacity.CurContentCG 'calc'd elsewhere
            Cwt = cComponent.Capacity.CurVol * dblDensity
        ElseIf dblDensity > 0.0000001 Then
            'calculate weight using density
            Ccg = cComponent.Capacity.CurContentCG 'calc'd elsewhere
            Cwt = cComponent.Capacity.CurVol * dblDensity
        ElseIf Var2Dbl(cTruck.Components(i%).Capacity.DefaultWtContents) > 0 Then
            'calculate using Default(user) content weight
            Ccg = cComponent.Capacity.CurContentCG 'calc'd elsewhere
            Cwt = cComponent.Capacity.CurVol * dblDensity
        Else
            'No contents
            Cwt = 0#
        End If
        
        'Ccg = Ccg + cTruck.Components(i%).Offset
        If cComponent.LocationReference = orBodyOrigin Then
            If cTruck.IsStandardMount Then
                cM = Cwt * (WB - L - Ccg)
            Else
                cM = Cwt * (WB - (L + BodyLength - Ccg))
            End If
        ElseIf cComponent.LocationReference = orFrontAxle Then
            cM = Cwt * (WB - Ccg)
        End If
        'Sum Moments
        dblSumMoments = dblSumMoments + (eM + cM)
    Next i%
    
    'Sum the Tag forces (negative) ----------
    lmt = cTruck.Chassis.Tags.Count
    For i% = 1 To lmt
        'Equipment
        Ccg = cTruck.Chassis.Tags(i%).Location
        Cwt = cTruck.Chassis.Tags(i%).Weight
        eM = Cwt * (WB - Ccg)
        'Applied force
        Cwt = -cTruck.Chassis.Tags(i%).DownwardForce
        cM = Cwt * (WB - Ccg)
        'Sum Moments
        dblSumMoments = dblSumMoments + (eM + cM)
    Next i%
    
    'Front Axle - Empty Chassis Weight ----------
    Ccg = 0
    Cwt = cTruck.Chassis.WtFront + cTruck.WtAdjustFront
    eM = Cwt * (WB - Ccg)
    'Sum Moments
    dblSumMoments = dblSumMoments + eM
    
    'Final Calculation ----------
    FrontLoading = dblSumMoments / WB

    Set cTank = Nothing
    Set cComponent = Nothing
End Function


Public Function RearLoading(cTruck As clsTruck) As Double
    'Standard Mount
    '--------------
    '   RearAxleLoad = sum[Cwt(l + Ccg)] / WB
    '       where Cwt = Component Weight
    '             Ccg = Dist from Component CG to Body Origin
    '               l = Dist from front axle to Body Origin
    '              WB = Dist between front and rear axle
    
    'Reverse Mount
    '-------------
    '   RearAxleLoad = sum[Cwt(l + BodyLength - Ccg)] / WB
    '       where Cwt = Component Weight
    '             Ccg = Dist from Component CG to Body Origin
    '               l = Dist from front axle to Body Origin
    '              WB = Dist between front and rear axle
    
    Dim dblSumMoments As Double
    Dim eM As Double 'Moment due to equipment weight
    Dim cM As Double 'Moment due to contents weight only
    Dim Cwt As Double
    Dim Ccg As Double
    Dim L As Double
    Dim WB As Double
    Dim BodyLength As Double
    Dim dblDensity As Double
    Dim i%
    Dim lmt As Integer
    Dim cTank As clsTank
    Dim cComponent As clsComponent

    BodyLength = cTruck.Body.BodyLength
    WB = cTruck.Chassis.WB
    L = cTruck.BodyLocation
    
    'Sum the Tank Loads -------------------
    'Contents
    lmt = cTruck.Body.Tanks.Count
    For i% = 1 To lmt
        Set cTank = cTruck.Body.Tanks(i%)
        dblDensity = cTank.DensityContents
        'calculate weight using density
        Ccg = cTank.CurContentCG 'calc'd elsewhere
        Cwt = cTank.CurVol * dblDensity
    
        If cTruck.IsStandardMount Then
            cM = Cwt * (L + Ccg)
        Else
            cM = Cwt * (L + BodyLength - Ccg)
        End If
        'Sum Moments
        dblSumMoments = dblSumMoments + cM
    Next i%
    'Equipment
    Ccg = cTruck.Body.EmptyCG
    Cwt = cTruck.Body.EmptyWeight
    If cTruck.IsStandardMount Then
        eM = Cwt * (L + Ccg)
    Else
        eM = Cwt * (L + BodyLength - Ccg)
    End If
    'Sum Moments
    dblSumMoments = dblSumMoments + eM
    
    'Sum the Component Loads --------------
    lmt = cTruck.Components.Count
    For i% = 1 To lmt
        Set cComponent = cTruck.Components(i%)
        'Equipment
        Cwt = cComponent.EmptyWeight
        Ccg = cComponent.EmptyCG + cComponent.Offset
        If cComponent.LocationReference = orBodyOrigin Then
            If cTruck.IsStandardMount Then
                eM = Cwt * (L + Ccg)
            Else
                eM = Cwt * (L + BodyLength - Ccg)
            End If
        ElseIf cComponent.LocationReference = orFrontAxle Then
                eM = Cwt * Ccg
        End If
        
        'Contents
        dblDensity = cComponent.Capacity.DensityContents
        If Var2Dbl(cComponent.Capacity.DefaultVolContents) > 0 Then
            'calculate using Default(user) content Volume
            Ccg = cComponent.Capacity.CurContentCG 'calc'd elsewhere
            'Cwt = cTruck.Components(i%).Capacity.DefaultVolContents * dblDensity
            Cwt = cComponent.Capacity.CurVol * dblDensity
        ElseIf dblDensity > 0.0000001 Then
            'calculate weight using density
            Ccg = cComponent.Capacity.CurContentCG 'calc'd elsewhere
            Cwt = cComponent.Capacity.CurVol * dblDensity
        ElseIf Var2Dbl(cComponent.Capacity.DefaultWtContents) > 0 Then
            'calculate using Default(user) content weight
            Ccg = cComponent.Capacity.CurContentCG 'calc'd elsewhere
            'Cwt = cTruck.Components(i%).Capacity.DefaultWtContents
            Cwt = cComponent.Capacity.CurVol * dblDensity
        Else
            'No contents
            Cwt = 0#
        End If
        
        'Ccg = Ccg + cTruck.Components(i%).Offset
        If cComponent.LocationReference = orBodyOrigin Then
            If cTruck.IsStandardMount Then
                cM = Cwt * (L + Ccg)
            Else
                cM = Cwt * (L + BodyLength - Ccg)
            End If
        ElseIf cComponent.LocationReference = orFrontAxle Then
            cM = Cwt * Ccg
        End If
        'Sum Moments
        dblSumMoments = dblSumMoments + (eM + cM)
    Next i%
    
    'Sum the Tag forces (negative) ----------
    lmt = cTruck.Chassis.Tags.Count
    For i% = 1 To lmt
        'Equipment
        Ccg = cTruck.Chassis.Tags(i%).Location
        Cwt = cTruck.Chassis.Tags(i%).Weight
        eM = Cwt * Ccg
        'Applied force
        Cwt = -cTruck.Chassis.Tags(i%).DownwardForce
        cM = Cwt * Ccg
        'Sum Moments
        dblSumMoments = dblSumMoments + (eM + cM)
    Next i%
    
    'Rear Axle - Empty Chassis Weight ----------
    Ccg = cTruck.Chassis.WB
    Cwt = cTruck.Chassis.WtRear + cTruck.WtAdjustRear
    eM = Cwt * Ccg
    'Sum Moments
    dblSumMoments = dblSumMoments + eM
    
    'Final Calculation ----------
    RearLoading = dblSumMoments / WB

    Set cTank = Nothing
    Set cComponent = Nothing
End Function


Public Function LoadingViolations(cTruck As clsTruck, IsOnRoad As Boolean) As String
    'Returns a string to display in a message box that indicates what
    ' cause the bridge law to be violated.  The function also looks
    ' for under-loading (uses MinWtFront and MinWtRear) and
    ' over-loading (based on mfg limits).
    ' If false, the 'IsOnRoad' flag will turn off evaluation of the US Bridge Laws.
    ' If no violations or other loading issues, an empty string is returned.
    Dim cAxle As clsAxleGroup
    Dim cAxles As clsAxleGroups
    Dim i%
    Dim dblTotal As Double
    Dim dblExcess As Double
    Dim sExcess As String
    Dim sRtn As String
    Dim sDescr As String
    
    sRtn = "" 'default to no problems
    
    'Define Axles (in order) -----------------
    Set cAxles = LoadedAxles(cTruck)
    
    'Check Gross Wt (Mfg Limit) -----------------
    dblTotal = 0#
    For i% = 1 To cAxles.Count
        dblTotal = dblTotal + cAxles(i%).ActualLd
    Next i%
    If dblTotal > cTruck.Chassis.WtLimitTotal Then
        dblExcess = (dblTotal - cTruck.Chassis.WtLimitTotal) * cGlobalInfo.MassUnits.Multiplier
        sExcess = Format(dblExcess, "### ") & cGlobalInfo.MassUnits.Display
        sRtn = "Exceeded Mfg Gross Wt Limit by " & sExcess & vbCrLf
    End If
    
    'Check Individual Axles (Mfg Limit & Min limit) -----------------
    For i% = 1 To cAxles.Count
        Set cAxle = cAxles(i%)
        If i% = 2 And cTruck.Chassis.TwinSteerSeparation > 0 Then
            'Skip second steer axle since both fail if one fails
        Else
            If cAxle.ActualLd > cAxle.AllowableLd Then
                If cTruck.Chassis.TwinSteerSeparation > 0 Then
                    'Dual-steer axles, so double the numbers
                    sDescr = "Front Steer Tandem"
                    dblExcess = (cAxle.ActualLd - cAxle.AllowableLd) _
                                * cGlobalInfo.MassUnits.Multiplier * 2
                Else
                    sDescr = cAxle.sDescription
                    dblExcess = (cAxle.ActualLd - cAxle.AllowableLd) _
                                * cGlobalInfo.MassUnits.Multiplier
                End If
                sExcess = Format(dblExcess, "### ") & cGlobalInfo.MassUnits.Display
                sRtn = sRtn & "Exceeded Mfg Wt Limit for " & sDescr & _
                       " by " & sExcess & vbCrLf
            End If
            If cAxle.ActualLd < cAxle.MinLoading Then
                If cTruck.Chassis.TwinSteerSeparation > 0 Then
                    'Dual-steer axles, so double the numbers
                    sDescr = "Front Steer Tandem"
                    dblExcess = cAxle.ActualLd * cGlobalInfo.MassUnits.Multiplier * 2
                Else
                    sDescr = cAxle.sDescription
                    dblExcess = cAxle.ActualLd * cGlobalInfo.MassUnits.Multiplier
                End If
                sExcess = Format(dblExcess, "### ") & cGlobalInfo.MassUnits.Display
                sRtn = sRtn & cAxle.sDescription & " was too light at " & sExcess & vbCrLf
            End If
        End If
    Next i%
    
    'Check Bridge Law -----------------
    If IsOnRoad Then
        'Evaluate Bridge Law for on-road use
        sRtn = sRtn & BridgeLawWarnings(cTruck.BridgeLaw, cAxles)
        'Check for over-loaded tanks
        sRtn = sRtn & TankOverloading(cTruck)
    End If
    
    LoadingViolations = sRtn
    
CleanUp:
    Set cAxle = Nothing
    Set cAxles = Nothing
End Function


Private Function TankOverloading(cTruck As clsTruck) As String
    'This routine is called (only) to see if any tanks are over-loaded.
    ' This function ONLY:
    '  - is called from 'LoadingViolations' function
    '  - applies to on-road situations
    '  - applies to Emusion-carrying tanks (other tanks chosen specifically for use)
    Dim i%
    Dim dblCurMass As Double
    Dim dblExcess As Double
    Dim sExcess As String
    Dim sRtn As String
    Dim cTank As clsTank
    
    For i% = 1 To cTruck.Body.Tanks.Count
        Set cTank = cTruck.Body.Tanks(i%)
        If cTank.CurTankUse = ttEmulsion _
        And cTank.MaxMass <> 0 Then
            'Check mass loading of this tank since it contains emulsion and
            ' has a defined mass limit
            dblCurMass = cTank.CurVol * cTank.DensityContents
            If dblCurMass > cTank.MaxMass Then
                dblExcess = (dblCurMass - cTank.MaxMass) * cGlobalInfo.MassUnits.Multiplier
                sExcess = Format(dblExcess, "### ") & cGlobalInfo.MassUnits.Display
                sRtn = sRtn & "Tank #" & CStr(i%) & " exceeded '" & cTank.MaxMassDesc & _
                       "' by " & sExcess & vbCrLf
            End If
        End If
    Next i%
    TankOverloading = sRtn
    Set cTank = Nothing
End Function

Public Function BridgeLawWarnings(BridgeLaw As String, cAxles As clsAxleGroups, Optional dblGrossAllowable As Double) As String
    ' The optional variable 'dblGrossAllowable' returns the Bridge Law limit for
    ' ONLY the parent group (not subgroups).
    Select Case UCase$(BridgeLaw)
    Case "US"
        BridgeLawWarnings = USBridgeLawWarnings(cAxles, dblGrossAllowable)
    Case Else
        'Other contries to go here eventually
    End Select
End Function

Public Function USBridgeLawWarnings(cAxles As clsAxleGroups, Optional dblGrossAllowable As Double) As String
    'Checks a group of axles to see if there are any US Bridge Law
    ' violations for the group or subgroup thereof
    ' The optional variable 'dblGrossAllowable' returns the Bridge Law limit for
    ' ONLY the parent group (not subgroups).
    Dim cAxleGroup As clsAxleGroup
    Dim cAxleGroups As clsAxleGroups
    Dim i%
    Dim x%
    Dim sRtn As String
    Dim dblExcess As Double
    Dim sExcess As String
    Dim iGrps As Integer
    Dim sGroup As String
    Dim sDesc As String
    Dim lAxles As Long
    Dim lSpc As Long
    Dim Ld As Double  'Axle Group Loading
    Dim iA As Integer 'First axle in group
    Dim iB As Integer 'Last axle in group
        
    sRtn = "" 'Default to no error
    Set cAxleGroups = New clsAxleGroups
    'Build AxleGroups
    Select Case cAxles.Count
    Case 2
        iGrps = 1
    Case 3
        iGrps = 3 '2+1
    Case 4
        iGrps = 6 '3+2+1
    Case 5
        iGrps = 10 '4+3+2+1
    Case 6
        iGrps = 15 '5+4+3+2+1
    Case 7
        iGrps = 21 '6+5+4+3+2+1
    Case 8
        iGrps = 28 '7+6+5+4+3+2+1
    End Select
    
    'Create Axle Groups -------
    iA = 1
    iB = 1
    For i% = 1 To iGrps
        iB = iB + 1
        If iB > cAxles.Count Then
            iA = iA + 1
            iB = iA + 1
        End If
        'Create Group name and descriptions & calc total loading for group
        sGroup = CStr(iA)
        sDesc = cAxles(iA).sDescription
        Ld = cAxles(iA).ActualLd
        lAxles = 1
        For x% = iA + 1 To iB
            sGroup = sGroup & "-" & CStr(x%)
            sDesc = sDesc & "-" & cAxles(x%).sDescription
            Ld = Ld + cAxles(x%).ActualLd
            lAxles = lAxles + 1
        Next x%
        Set cAxleGroup = New clsAxleGroup
        cAxleGroup.sGroup = sGroup
        cAxleGroup.sDescription = sDesc
        'Spacing saved in [ft] and rounded to the nearest foot
        lSpc = CLng(Round(3.2808398950131 * (cAxles(iB).AxleLoc - cAxles(iA).AxleLoc), 0))
        cAxleGroup.Spacing = lSpc
        cAxleGroup.ActualLd = Ld
        cAxleGroup.NumAxles = lAxles
        cAxleGroups.Add cAxleGroup, cAxleGroup.sGroup
    Next i%
    
    'Evaluate Groups
    For i% = 1 To iGrps
        'Calculate allowable load for group
        lAxles = cAxleGroups(i%).NumAxles
        lSpc = cAxleGroups(i%).Spacing
        cAxleGroups(i%).AllowableLd = 500 * Round(((lSpc * lAxles) / (lAxles - 1) + 12 * lAxles + 36), 0)
        'Make sure axles aren't more than 20,000lb each
        If cAxleGroups(i%).AllowableLd > (lAxles * 20000) Then
            cAxleGroups(i%).AllowableLd = lAxles * 20000
        End If
        
        cAxleGroups(i%).AllowableLd = 0.45359237 * cAxleGroups(i%).AllowableLd 'lb to kg
        'Save Gross Allowable Weight (all axles)
        If i% = (cAxles.Count - 1) Then
            dblGrossAllowable = cAxleGroups(i%).AllowableLd
        End If
        'Create warnings for violations
        If cAxleGroups(i%).ActualLd > cAxleGroups(i%).AllowableLd Then
            dblExcess = (cAxleGroups(i%).ActualLd - cAxleGroups(i%).AllowableLd) * cGlobalInfo.MassUnits.Multiplier
            sExcess = Format(dblExcess, "### ") & cGlobalInfo.MassUnits.Display
            sRtn = sRtn & "Exceeded Bridge Law Limit for the following Axle-Group: " & cAxleGroups(i%).sDescription & _
                   " by " & sExcess & vbCrLf
        End If
    Next i%
    USBridgeLawWarnings = sRtn

CleanUp:
    Set cAxleGroup = Nothing
    Set cAxleGroups = Nothing
End Function


Public Function LoadingSummary(cTruck As clsTruck, IsOnRoad As Boolean) As clsAxleGroups
    'Returns group of axles to be used for loading summary report
    ' the .AllowableLd property of all returned "axles" is the Mfg Limit with the
    ' exception of the GVWA which is BridgeLaw limit instead of MfgLimit _IF_ the
    ' IsOnRoad flag is set.
    Dim cAxle As clsAxleGroup
    Dim cAllAxles As clsAxleGroup
    Dim cAxles As clsAxleGroups
    Dim cAxlesRtn As clsAxleGroups
    Dim cTempGroup As clsAxleGroups
    Dim dblTruckWeight As Double
    Dim i%
    Dim intAxle As Integer
    Dim dblGrossAllowable As Double
    
    Dim sRtn As String
    
    'Define Axles (in order) -----------------
    Set cAxles = LoadedAxles(cTruck)
    
    'Create an axle-group for the GVW
    Set cAllAxles = New clsAxleGroup
    cAllAxles.sDescription = "GVW"
    cAllAxles.NumAxles = intAxle
    cAllAxles.Spacing = cAxles(cAxles.Count).AxleLoc - cAxles(1).AxleLoc
    cAllAxles.AllowableLd = cTruck.Chassis.WtLimitTotal 'Start w/ Mfg. limit
    If cAllAxles.AllowableLd = 0 Then
        cAllAxles.AllowableLd = cTruck.Chassis.WtLimitFront + cTruck.Chassis.WtLimitRear
    End If
    dblTruckWeight = 0#
    For i% = 1 To cAxles.Count
        dblTruckWeight = dblTruckWeight + cAxles(i%).ActualLd
    Next i%
    cAllAxles.ActualLd = dblTruckWeight
    
    'Find Bridge Law Limit for GVW -----------------
    If IsOnRoad Then 'Evaluate Bridge Law for on-road use (only care about GVWA)
        'Check Gross Wt (Mfg Limit if more restrictive) -----------------
        sRtn = BridgeLawWarnings(cTruck.BridgeLaw, cAxles, dblGrossAllowable)
        If cAllAxles.AllowableLd = 0 Or dblGrossAllowable < cAllAxles.AllowableLd Then
            cAllAxles.AllowableLd = dblGrossAllowable
        End If
    End If
    
    'Add the whole truck as one of the axlegroups
    cAxles.Add cAllAxles
    
    'Return cAxles (group Tandem axles)
    Set cAxlesRtn = New clsAxleGroups
    For i% = 1 To cAxles.Count
        If Left$(cAxles(i%).sDescription, 6) = "Tandem" Then
            If cAxles(i%).sDescription = "Tandem Front" Then
                'Return only a single "axle" for the rear tandem
                Set cAxle = cAxles(i%)
                cAxle.sDescription = "Rear"
                cAxle.AllowableLd = cAxles("Tandem Front").AllowableLd + cAxles("Tandem Rear").AllowableLd
                cAxle.ActualLd = cAxles("Tandem Front").ActualLd + cAxles("Tandem Rear").ActualLd
                cAxle.MinLoading = cAxles("Tandem Front").MinLoading + cAxles("Tandem Rear").MinLoading
                cAxle.AxleLoc = (cAxles("Tandem Front").AxleLoc + cAxles("Tandem Rear").AxleLoc) / 2
                cAxlesRtn.Add cAxle, cAxle.sDescription
            End If
        ElseIf Right$(cAxles(i%).sDescription, 10) = "Steer Axle" Then
            If cAxles(i%).sDescription = "First Steer Axle" Then
                'Return only a single "axle" for the steering tandem
                Set cAxle = cAxles(i%)
                cAxle.sDescription = "Front"
                cAxle.AllowableLd = cAxles(1).AllowableLd * 2
                cAxle.ActualLd = cAxles(1).ActualLd * 2
                cAxle.MinLoading = cAxles(1).MinLoading * 2
                cAxle.AxleLoc = 0
                cAxlesRtn.Add cAxle, cAxle.sDescription
            End If
        Else
            cAxlesRtn.Add cAxles(i%), cAxles(i%).sDescription
        End If
    Next i%
    Set LoadingSummary = cAxlesRtn
    
CleanUp:
    Set cAxle = Nothing
    Set cAxles = Nothing
    Set cAxlesRtn = Nothing
    Set cAllAxles = Nothing
    Set cTempGroup = Nothing
End Function


Public Function LoadedAxles(cTruck As clsTruck) As clsAxleGroups
    'Returns AxleGroups of all axles that are loaded (i.e. - NOT retracted tags)
    Dim i%
    Dim intTemp As Integer
    Dim intNumTags As Integer
    Dim intAxle As Integer
    Dim intTags As Integer
    Dim cAxle As clsAxleGroup
    Dim cAxles As clsAxleGroups
    Dim MinWtFront As Double
    Dim MinWtRear As Double
    
    'Minimum allowable weight same as chassis weight
    MinWtFront = cTruck.Chassis.WtFront
    MinWtRear = cTruck.Chassis.WtRear
    
    'Define Axles (in order) -----------------
    Set cAxle = New clsAxleGroup
    Set cAxles = New clsAxleGroups
    intAxle = 1
    
    ' Front Axle
    If cTruck.Chassis.TwinSteerSeparation = 0 Then
        'Single front axle
        cAxle.sDescription = "Front"
        cAxle.sGroup = CStr(intAxle)
        cAxle.AllowableLd = cTruck.Chassis.WtLimitFront
        cAxle.ActualLd = FrontLoading(cTruck)
        cAxle.AxleLoc = 0
        cAxle.MinLoading = MinWtFront
        cAxle.ForceToPressure = 0 'N/A
        cAxles.Add cAxle, cAxle.sDescription
    Else
        'Dual front axles
        cAxle.sDescription = "First Steer Axle"
        cAxle.sGroup = CStr(intAxle)
        cAxle.AllowableLd = cTruck.Chassis.WtLimitFront / 2
        cAxle.ActualLd = FrontLoading(cTruck) / 2
        cAxle.AxleLoc = -cTruck.Chassis.TwinSteerSeparation / 2
        cAxle.MinLoading = MinWtFront / 2
        cAxle.ForceToPressure = 0 'N/A
        cAxles.Add cAxle, cAxle.sDescription
        '-
        intAxle = intAxle + 1
        Set cAxle = New clsAxleGroup
        cAxle.sDescription = "Second Steer Axle"
        cAxle.sGroup = CStr(intAxle)
        cAxle.AllowableLd = cTruck.Chassis.WtLimitFront / 2
        cAxle.ActualLd = FrontLoading(cTruck) / 2
        cAxle.AxleLoc = cTruck.Chassis.TwinSteerSeparation / 2
        cAxle.MinLoading = MinWtFront / 2
        cAxle.ForceToPressure = 0 'N/A
        cAxles.Add cAxle, cAxle.sDescription
    End If
    
    ' Count the # of intermediate tags (aka "pushers")
    intTags = cTruck.Chassis.Tags.Count
    intNumTags = 0
    For i% = 1 To intTags
        If cTruck.Chassis.Tags(i%).Location < cTruck.Chassis.WB _
         And cTruck.Chassis.Tags(i%).DownwardForce > 0 Then
            'This tag is in front of the rear axle
            intNumTags = intNumTags + 1
        End If
    Next i%
    
    ' Intermediate tags (if any)
    intTemp = 0
    For i% = 1 To intTags
        If cTruck.Chassis.Tags(i%).Location < cTruck.Chassis.WB _
         And cTruck.Chassis.Tags(i%).DownwardForce > 0 Then
            'This tag is in front of the rear axle
            Set cAxle = New clsAxleGroup
            If intNumTags > 1 Then
                intTemp = intTemp + 1
                cAxle.sDescription = "Pusher " & intTemp
            Else
                cAxle.sDescription = "Pusher" 'Only one Pusher, so no index suffix
            End If
            intAxle = intAxle + 1
            cAxle.sGroup = CStr(intAxle)
            cAxle.AllowableLd = cTruck.Chassis.Tags(i%).WtLimit
            cAxle.ActualLd = cTruck.Chassis.Tags(i%).DownwardForce
            cAxle.AxleLoc = cTruck.Chassis.Tags(i%).Location
            cAxle.MinLoading = 0
            cAxle.ForceToPressure = cTruck.Chassis.Tags(i%).ForceToPressure
            cAxles.Add cAxle, cAxle.sDescription
        End If
    Next
    
    ' Rear Axle(s)
    If cTruck.Chassis.TandemSpacing = 0 Then
        'Single axle Rear
        Set cAxle = New clsAxleGroup
        cAxle.sDescription = "Rear"
        intAxle = intAxle + 1
        cAxle.sGroup = CStr(intAxle)
        cAxle.AllowableLd = cTruck.Chassis.WtLimitRear
        cAxle.ActualLd = RearLoading(cTruck)
        cAxle.AxleLoc = cTruck.Chassis.WB
        cAxle.MinLoading = MinWtRear
        cAxle.ForceToPressure = 0 'N/A
        cAxles.Add cAxle, cAxle.sDescription
    Else 'Rear Tandem (2 axles)
        'Front-most tandem axle
        Set cAxle = New clsAxleGroup
        cAxle.sDescription = "Tandem Front"
        intAxle = intAxle + 1
        cAxle.sGroup = CStr(intAxle)
        cAxle.AllowableLd = cTruck.Chassis.WtLimitRear / 2
        cAxle.ActualLd = RearLoading(cTruck) / 2
        cAxle.AxleLoc = cTruck.Chassis.WB - (cTruck.Chassis.TandemSpacing / 2)
        cAxle.MinLoading = MinWtRear / 2
        cAxle.ForceToPressure = 0 'N/A
        cAxles.Add cAxle, cAxle.sDescription
        'Rear-most tandem axle
        intAxle = intAxle + 1
        Set cAxle = New clsAxleGroup
        cAxle.sDescription = "Tandem Rear"
        cAxle.sGroup = CStr(intAxle)
        cAxle.AllowableLd = cTruck.Chassis.WtLimitRear / 2
        cAxle.ActualLd = RearLoading(cTruck) / 2
        cAxle.AxleLoc = cTruck.Chassis.WB + (cTruck.Chassis.TandemSpacing / 2)
        cAxle.MinLoading = MinWtRear / 2
        cAxle.ForceToPressure = 0 'N/A
        cAxles.Add cAxle, cAxle.sDescription
    End If
    
    ' Count the # of rear tags (aka "tags")
    intNumTags = 0
    For i% = 1 To intTags
        If cTruck.Chassis.Tags(i%).Location > cTruck.Chassis.WB _
         And cTruck.Chassis.Tags(i%).DownwardForce > 0 Then
            'This tag is in front of the rear axle
            intNumTags = intNumTags + 1
        End If
    Next i%
    
    ' Rear tags (if any)
    intTemp = 0
    For i% = 1 To intTags
        If cTruck.Chassis.Tags(i%).Location > cTruck.Chassis.WB _
         And cTruck.Chassis.Tags(i%).DownwardForce > 0 Then
            'This tag is in front of the rear axle
            Set cAxle = New clsAxleGroup
            If intNumTags > 1 Then
                intTemp = intTemp + 1
                cAxle.sDescription = "Tag " & intTemp
            Else
                cAxle.sDescription = "Tag" 'Only one Pusher, so no index suffix
            End If
            intAxle = intAxle + 1
            cAxle.sGroup = CStr(intAxle)
            cAxle.AllowableLd = cTruck.Chassis.Tags(i%).WtLimit
            cAxle.ActualLd = cTruck.Chassis.Tags(i%).DownwardForce
            cAxle.AxleLoc = cTruck.Chassis.Tags(i%).Location
            cAxle.MinLoading = 0
            cAxle.ForceToPressure = cTruck.Chassis.Tags(i%).ForceToPressure
            cAxles.Add cAxle, cAxle.sDescription
        End If
    Next i%

    Set LoadedAxles = cAxles
    Set cAxle = Nothing
    Set cAxles = Nothing
End Function
