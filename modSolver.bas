Attribute VB_Name = "modSolver"
Option Explicit
Public sngTimeBefore As Single

Public bIsOnRoad As Boolean
Public dblEmulPct As Double 'target Emulsion percentage of ANFO/Emulsion blend (0-100)
Public dblDFOPct As Double 'target percent DFO (0 if no FO tank, 0.06 otherwise)
Public dblPctComplete As Double ' updated for progress indicator
Public bTruckFileDirty As Boolean

Private MinWtFront As Double
Private MinWtRear As Double
Private colConfigs As Collection 'set of configurations for Dual-Use tanks.
                                 'Eval as binary where bit 0 is 1st dual-use tanks
                                 'and bit 1 is second dual-use tank

Private Const LITTLE_STEP = 1# / 50#
Private Const BIG_STEP = 1# / 10#


'NOTE
'Solver loop must calculate Sticklength, CG, and weight before calling
'axleload routines.

Public Sub SolveLoading(cTruck As clsTruck, Optional bFindBest As Boolean = False, Optional ProductLimit As Double = 0)
    'If bFindBest, then all tank configurations are considered.  Otherwise,
    ' only the current tank config is considered.
    Dim i%
    Dim sRtn As String
    Dim sMsg As String
    Dim dblDecrement As Double 'amt to reduce total product on each iteration
    Dim dblDecr_AN As Double 'amt to reduce AN on each iteration
    Dim dblDecr_Emul As Double 'amt to reduce Emul on each iteration
    Dim strResults As String
    Dim dblStepFraction As Double
    Dim bFoundSoln As Boolean
    Dim bFirstStep As Boolean
    Dim dblPctEmulTarget As Double
    Dim dblVal As Double
    Dim cBestLoadingCase As clsLoadingCase
    Dim cTank As clsTank
    On Error GoTo errHandler
    
    InitializeTanks cTruck
    BuildConfigurations cTruck, bFindBest
    Set cBestLoadingCase = New clsLoadingCase
    cBestLoadingCase.InitTanks cTruck
    cBestLoadingCase.TotalProductMass = 0
    
    'Check all possible configurations
    For i% = 1 To colConfigs.Count
        
        sRtn = SetTruckConfig(cTruck, colConfigs.Item(i%)) 'Configure tanks
        dblPctEmulTarget = AdaptiveEmulPct(cTruck)
        sRtn = DistributeProduct(cTruck, True, ProductLimit) 'Maximize product
        bFirstStep = True
        dblStepFraction = BIG_STEP
    
        'Set the step (decrement) amounts
        dblDecrement = (cTruck.Body.MassANTotal + cTruck.Body.MassEmulTotal) * dblStepFraction
        dblDecr_AN = dblDecrement * (100# - dblPctEmulTarget) / 100#
        dblDecr_Emul = dblDecrement * dblPctEmulTarget / 100#
        
        Do
            sRtn = DistributeProduct(cTruck) 'distr. w\ use current amounts
            bFoundSoln = False
            sRtn = FindDistribution(cTruck, dblStepFraction)
            If sRtn = "" Then
                bFoundSoln = True
            End If
            
            If bFoundSoln And dblStepFraction = BIG_STEP Then
                'Coarse step just yielded a success, and use smaller steps
                If Not bFirstStep Then
                    'Step back one first
                    'Increase product (to previous level)
                    cTruck.Body.MassANTotal = cTruck.Body.MassANTotal + dblDecr_AN
                    cTruck.Body.MassEmulTotal = cTruck.Body.MassEmulTotal + dblDecr_Emul
                End If
                'Set the finer step size
                dblStepFraction = LITTLE_STEP
                dblDecrement = (cTruck.Body.MassANTotal + cTruck.Body.MassEmulTotal) * dblStepFraction
                dblDecr_AN = dblDecrement * (100# - dblPctEmulTarget) / 100#
                dblDecr_Emul = dblDecrement * dblPctEmulTarget / 100#
            ElseIf bFoundSoln Then
                'Found a solution at minimal step size
                Exit Do 'Exit with solution
            Else
                'Reduce product loading & Distribute
                cTruck.Body.MassANTotal = cTruck.Body.MassANTotal - dblDecr_AN
                If cTruck.Body.MassANTotal < 0# Then cTruck.Body.MassANTotal = 0#
                
                cTruck.Body.MassEmulTotal = cTruck.Body.MassEmulTotal - dblDecr_Emul
                If cTruck.Body.MassEmulTotal < 0# Then cTruck.Body.MassEmulTotal = 0#
            End If
            'Distribute product w\ use current amounts
            sRtn = DistributeProduct(cTruck)
            bFirstStep = False
        Loop Until (cTruck.Body.MassANTotal + cTruck.Body.MassEmulTotal) < (2 * dblDecrement)
        
        If bFoundSoln Then
            'See if this is best selection yet
            HighestCapacityLoading cTruck, cBestLoadingCase

'           'Add to collection of Good Truck configs
'           strResults = strResults & SummaryText(cTruck)
            
'            'show solution
'            frmMonitor.MonitorDwg cTruck, "No warnings." 'Graphical representation of loading"
'            Sleep 5000
        End If
    Next i%
    
'If strResults = "" Then strResults = "No solutions!"
''frmResults.ShowResults strResults
    
    If cBestLoadingCase.TotalProductMass = 0 Then
        'No Solution was found
        MsgBox "No loading solution was found.  Please review and change the options and try again.", vbExclamation
    Else
        'Set truck to best configuration found
        For i% = 1 To cTruck.Body.Tanks.Count
            Set cTank = cTruck.Body.Tanks(i%)
            cTank.CurVol = cBestLoadingCase.ContentVol(i%)
            cTank.CurTankUse = cBestLoadingCase.TankConfig(CStr(i%))
        Next i%
        SetProductDensities cTruck 'set product density based on updated CurTankUse
        ReCalc cTruck 'recaclc tank levels, stick height, fill relationships, etc.
        If bFindBest Then
            sMsg = "This is the optimized loading solution given the inputs provided."
        Else
            sMsg = "This is the optimized loading solution for the current tank configuration and the inputs provided."
        End If
        If ProductLimit > 0# Then
            dblVal = ProductLimit * cGlobalInfo.MassUnits.Multiplier
            sMsg = sMsg & vbCrLf & vbCrLf _
                   & "  NOTE: Solution considers a user-specified total product limit of " _
                   & Format(dblVal, "#,##0 ") & cGlobalInfo.MassUnits.Display & "."
        End If
        frmMonitor.MonitorDwg cTruck, sMsg 'Graphical representation of loading"
    End If

    Set cBestLoadingCase = Nothing
    Set cTank = Nothing
    Exit Sub
errHandler:
    ErrorIn "modSolver.SolveLoading(cTruck,bFindBest,ProductLimit)", Array(cTruck, bFindBest, _
         ProductLimit)
End Sub


Public Sub InitializeTanks(cTruck As clsTruck)
    'Calculates tank and component Maximum Heights
    ' This only needs to be done when loading a new truck file or
    ' when new tanks/components are added to an existing truck
    'Also set a default DFO Percent for use in calculations
    Dim i%, x%
    Dim lnum As Long
    Dim Ht As Double
    Dim colK As New Collection
    Dim lCount As Long
    On Error GoTo errHandler
    
    'Process Body tanks
    lnum = cTruck.Body.Tanks.Count
    For i% = 1 To lnum
        Set colK = cTruck.Body.Tanks(i%).StickLength
        Ht = SolvePolynomial(colK, 0)
        cTruck.Body.Tanks(i%).MaxHt = Ht
    Next i%
    'Process Components that are tanks
    lnum = cTruck.Components.Count
    For i% = 1 To lnum
        Err.Clear
        On Error Resume Next
        lCount = cTruck.Components(i%).Capacity.StickLength.Count
        If Err.Number = 0 Then
            If lCount > 0 Then
                'Only if Component has k-factors can a MaxHt be calculated
                Set colK = cTruck.Components(i%).Capacity.StickLength
                Ht = SolvePolynomial(colK, cTruck.Components(i%).Capacity.Volume)
                cTruck.Components(i%).Capacity.MaxHt = Ht
            End If
        End If
    Next i%
    Err.Clear
    On Error GoTo errHandler
    'Now set a value for %DFO based on whether truck has a DFO tank
    dblDFOPct = 0#
    lnum = cTruck.Components.Count
    For i% = 1 To lnum
        If cTruck.Components(i%).ContentsType = ctFuel Then
            For x% = 1 To cTruck.Components(i%).FillRelationShips.Count
                If cTruck.Components(i%).FillRelationShips(x%).ParentProduct = ptAN Then
                    dblDFOPct = cTruck.Components(i%).FillRelationShips(x%).Multiplier
                    Exit For
                End If
            Next
            Exit For
        End If
    Next i%
    
CleanUp:
    Set colK = Nothing
    Exit Sub
errHandler:
    ErrorIn "modSolver.InitializeTanks(cTruck)", cTruck
End Sub



Private Function BuildConfigurations(cTruck As clsTruck, bFindBest As Boolean)
    'Figures out all of the possible tank configurations, prioritizes them,
    ' and puts results in colConfigs.
    'Since there can only be up to 2 dual-use tanks on any body, there
    ' are a max of 4 possible configurations
    Dim i%
    Dim x%
    Dim lIndex As Long
    Dim intConfig As Integer
    Dim iNum As Integer
    Dim NumDualUse As Integer
    Dim NumConfigs As Integer
    Dim dblFullPct As Double
    Dim dblPctDiff As Double
    Dim dblANTot As Double
    Dim dblEmulTot As Double
    Dim dblSmallest As Double
    Dim Configs() As Double '    Configs(x,0) = Index, Configs(x,1) = PctDif
    On Error GoTo errHandler
    
    
    'First find out how many configurations there could be
    For i% = 1 To cTruck.Body.Tanks.Count
        If cTruck.Body.Tanks(i%).TankType = ttDual Then
            NumDualUse = NumDualUse + 1
        End If
    Next i%
    
    Set colConfigs = New Collection
    If NumDualUse = 0 Then
        'there's only one possible configuration
        colConfigs.Add 0, "1"
        Exit Function
    ElseIf bFindBest = False Then
        'Only return current configuration in this case
        intConfig = CurrentConfigVal(cTruck)
        colConfigs.Add intConfig, "1"
        Exit Function
    End If
    
    'See which configs are closest to ideal
    NumConfigs = 2 ^ (NumDualUse)
    ReDim Configs(NumConfigs - 1, 1)
    For i% = 0 To (NumConfigs - 1)
        Call SetTruckConfig(cTruck, i%)
        dblANTot = 0#
        dblEmulTot = 0#
        'sum components assuming full tanks
        For x% = 1 To cTruck.Body.Tanks.Count
            If cTruck.Body.Tanks(x%).CurTankUse = ttEmulsion Then
                'Increment Emulsion total
                dblEmulTot = dblEmulTot + cTruck.Body.Tanks(x%).Volume * cGlobalInfo.DensityEmul
            Else
                'Increment AN total
                dblANTot = dblANTot + cTruck.Body.Tanks(x%).Volume * cGlobalInfo.DensityAN
            End If
        Next x%
        'Calculate %Emul (taking DFO into acct!)
        dblFullPct = Round(100# * dblEmulTot / (dblEmulTot + dblANTot / (1# - dblDFOPct)), 1)
        If dblEmulPct > 0# And dblFullPct = 0# _
         Or dblEmulPct = 0 And dblFullPct > 0# Then
            'Do not add this configuration (i.e., no use having emul tank for 0% Blend)
        Else
            'add this config since it could work
            dblPctDiff = Abs(dblFullPct - dblEmulPct)
            Configs(lIndex, 0) = i%
            Configs(lIndex, 1) = dblPctDiff
            lIndex = lIndex + 1
        End If
    Next i%
    
    'Now create configuration collection starting with best combinations
    dblSmallest = 1E+100
    iNum = 0
    Do
        For i% = 0 To lIndex - 1
            If Configs(i%, 1) < dblSmallest Then
                dblSmallest = Configs(i%, 1)
            End If
        Next i%
        For i% = 0 To lIndex - 1
            If Configs(i%, 1) = dblSmallest Then
                iNum = iNum + 1
                colConfigs.Add Configs(i%, 0), CStr(iNum)
                'now reset the the value so it won't be found on the next scan
                Configs(i%, 1) = 9E+100
                dblSmallest = 1E+100
            End If
        Next i%
    Loop Until iNum = lIndex
    Exit Function
errHandler:
    ErrorIn "modSolver.BuildConfigurations(cTruck,bFindBest)", Array(cTruck, bFindBest)
End Function


Public Function SetTruckConfig(cTruck As clsTruck, intConfig As Integer)
    'For the Truck referenced, configures the dual-use tanks by intConfig:
    '  0 = All dual-use tank(s) > AN
    '  1 = First dual-use tank > EMUL, (if appl) Second dual-use tank > AN
    '  2 = First dual-use tank > AN, Second dual-use tank > EMUL
    '  3 = Both dual-use tanks > EMUL
    Dim i%
    Dim bFirstSet As Boolean
    Dim UseFlag As TankUse '0=AN, 1=Emul
    On Error GoTo errHandler
    
    'Loop through tanks to fink dual-use tanks
    For i% = 1 To cTruck.Body.Tanks.Count
        If cTruck.Body.Tanks(i%).TankType = ttDual Then
            If Not bFirstSet Then
                UseFlag = 1 And intConfig
                bFirstSet = True
            Else
                UseFlag = (2 And intConfig) / 2
            End If
            cTruck.Body.Tanks(i%).CurTankUse = UseFlag
        End If
    Next i%
    SetProductDensities cTruck
    Exit Function
errHandler:
    ErrorIn "modSolver.SetTruckConfig(cTruck,intConfig)", Array(cTruck, intConfig)
End Function


Private Function DistributeProduct(cTruck As clsTruck, Optional bMaximize As Boolean = False, Optional ProductLimit As Double = 0) As String
    'Uses dblEmulPct and distributes product in tanks for current configuration;
    'loading the rear-most compartments first.  No attempt is made here to see
    ' if the truck may be overloaded.
    Dim i%
    Dim sRtn As String
    Dim dblANTot As Double
    Dim dblEmulTot As Double
    Dim dblTemp As Double
    Dim dblPctAct As Double
    Dim dblPctEmulTarget As Double 'Desired Emul% for >>this configuration<<
    Dim dblTarget As Double
    Dim dblReduce As Double
    Dim intTemp As Integer
    Dim cEmulTanks As clsTanks
    Dim cANTanks As clsTanks
    On Error GoTo errHandler
    
    DistributeProduct = "" 'default to OK
        
    'Change target Emul% for this truck  if necessary
    For i% = 1 To cTruck.Body.Tanks.Count
        intTemp = intTemp + cTruck.Body.Tanks(i%).CurTankUse
    Next i%
    dblPctEmulTarget = AdaptiveEmulPct(cTruck)
    
    If bMaximize Then
        '--- sum components (setting tanks full)
        For i% = 1 To cTruck.Body.Tanks.Count
            cTruck.Body.Tanks(i%).CurVol = cTruck.Body.Tanks(i%).Volume
            If cTruck.Body.Tanks(i%).CurTankUse = ttEmulsion Then
                'Increment Emulsion total
                dblEmulTot = dblEmulTot + cTruck.Body.Tanks(i%).CurVol * cTruck.Body.Tanks(i%).DensityContents
            Else
                'Increment AN total
                dblANTot = dblANTot + cTruck.Body.Tanks(i%).CurVol * cTruck.Body.Tanks(i%).DensityContents
            End If
        Next i%
    Else
        '--- use supplied amounts
        dblEmulTot = cTruck.Body.MassEmulTotal
        dblANTot = cTruck.Body.MassANTotal
        'Distribute dblEmulTot into Emulsion tanks
        Set cEmulTanks = New clsTanks
        Set cEmulTanks = EmulsionTanks(cTruck)
        dblTemp = dblEmulTot / cGlobalInfo.DensityEmul
        For i% = cEmulTanks.Count To 1 Step -1
            cEmulTanks(i%).CurVol = dblTemp
            If cEmulTanks(i%).CurVol > cEmulTanks(i%).Volume Then
                cEmulTanks(i%).CurVol = cEmulTanks(i%).Volume
            End If
            dblTemp = dblTemp - cEmulTanks(i%).CurVol
        Next i%
        'Distribute dblANTot into AN tanks
        Set cANTanks = New clsTanks
        Set cANTanks = ANTanks(cTruck)
        dblTemp = dblANTot / cGlobalInfo.DensityAN
        For i% = cANTanks.Count To 1 Step -1
            cANTanks(i%).CurVol = dblTemp
            If cANTanks(i%).CurVol > cANTanks(i%).Volume Then
                cANTanks(i%).CurVol = cANTanks(i%).Volume
            End If
            dblTemp = dblTemp - cANTanks(i%).CurVol
        Next i%
    End If
    
    'Calculate actual %Emul (taking DFO into acct!)
    dblPctAct = 100# * dblEmulTot / (dblEmulTot + dblANTot / (1# - dblDFOPct))
    '--- Modify (total)product to achieve desired %Emul
    If dblPctAct > dblPctEmulTarget Then
        'Too much emulsion, Calculate Emul based on ANFO
        'E = ANFO/[%E-1] = AN / [ (1-%DFO)(100/PE-1) ]
        dblTarget = dblANTot / ((1# - dblDFOPct) * (100 / dblPctEmulTarget - 1))
        dblReduce = dblEmulTot - dblTarget 'need to reduce Emul by this amt
        'Reduce Emul in tanks starting in 1st tank
        Set cEmulTanks = New clsTanks
        Set cEmulTanks = EmulsionTanks(cTruck)
        For i% = 1 To cEmulTanks.Count
            If (cEmulTanks(i%).CurVol * cEmulTanks(i%).DensityContents) >= dblReduce Then
                'dblReduce can be taken out of this tank alone
                cEmulTanks(i%).CurVol = cEmulTanks(i%).CurVol _
                                        - (dblReduce / cEmulTanks(i%).DensityContents)
                dblReduce = 0
                Exit For
            Else
                'This tank will be set empty and some of another will have to be empties too
                dblReduce = dblReduce - (cEmulTanks(i%).CurVol * cEmulTanks(i%).DensityContents)
                cEmulTanks(i%).CurVol = 0
            End If
        Next i%
    ElseIf dblPctEmulTarget > dblPctAct Then
        'Too much ANFO, Calculate AN based on Emul
        'AN = E * [ (1-%DFO)(100/PE-1) ]
        dblTarget = dblEmulTot * (1# - dblDFOPct) * (100 / dblPctEmulTarget - 1)
        dblReduce = dblANTot - dblTarget 'need to reduce AN by this amt
        'Reduce AN in tanks starting in 1st tank
        Set cANTanks = New clsTanks
        Set cANTanks = ANTanks(cTruck)
        For i% = 1 To cANTanks.Count
            If (cANTanks(i%).CurVol * cANTanks(i%).DensityContents) >= dblReduce Then
                'dblReduce can be taken out of this tank alone
                cANTanks(i%).CurVol = cANTanks(i%).CurVol _
                                        - (dblReduce / cANTanks(i%).DensityContents)
                dblReduce = 0
                Exit For
            Else
                'This tank will be set empty and some of another will have to be empties too
                dblReduce = dblReduce - (cANTanks(i%).CurVol * cANTanks(i%).DensityContents)
                cANTanks(i%).CurVol = 0
            End If
        Next i%
    Else
        'Just redistribute product
    End If
    
    '--- Reduce total available product if user specified a Product Limit
    If ProductLimit > 0 Then
        UpdateMassTotals cTruck
        dblTemp = cTruck.Body.MassANTotal * (1 + dblDFOPct) + cTruck.Body.MassEmulTotal
        If ProductLimit < dblTemp Then
            'ProductLimit is less than current starting point, so reduce product
            dblTemp = dblTemp - ProductLimit '=Amount we have to remove from total product
            'Reduce Emulsion
            dblReduce = dblPctEmulTarget / 100# * dblTemp
            Set cEmulTanks = New clsTanks
            Set cEmulTanks = EmulsionTanks(cTruck)
            For i% = 1 To cEmulTanks.Count
                If (cEmulTanks(i%).CurVol * cEmulTanks(i%).DensityContents) >= dblReduce Then
                    'dblReduce can be taken out of this tank alone
                    cEmulTanks(i%).CurVol = cEmulTanks(i%).CurVol _
                                            - (dblReduce / cEmulTanks(i%).DensityContents)
                    dblReduce = 0
                    Exit For
                Else
                    'This tank will be set empty and some of another will have to be empties too
                    dblReduce = dblReduce - (cEmulTanks(i%).CurVol * cEmulTanks(i%).DensityContents)
                    cEmulTanks(i%).CurVol = 0
                End If
            Next i%
            'Reduce AN
            Set cANTanks = New clsTanks
            Set cANTanks = ANTanks(cTruck)
            dblReduce = ((100# - dblPctEmulTarget) / 100# * dblTemp) / (1 + dblDFOPct)
            For i% = 1 To cANTanks.Count
                If (cANTanks(i%).CurVol * cANTanks(i%).DensityContents) >= dblReduce Then
                    'dblReduce can be taken out of this tank alone
                    cANTanks(i%).CurVol = cANTanks(i%).CurVol _
                                            - (dblReduce / cANTanks(i%).DensityContents)
                    dblReduce = 0
                    Exit For
                Else
                    'This tank will be set empty and some of another will have to be empties too
                    dblReduce = dblReduce - (cANTanks(i%).CurVol * cANTanks(i%).DensityContents)
                    cANTanks(i%).CurVol = 0
                End If
            Next i%
        End If
    End If
    
    '--- Update product totals
    UpdateMassTotals cTruck
    
    sRtn = sRtn & ReCalc(cTruck)
    DistributeProduct = sRtn
CleanUp:
    Set cEmulTanks = Nothing
    Set cANTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.DistributeProduct(cTruck,bMaximize,ProductLimit)", Array(cTruck, bMaximize, _
         ProductLimit)
    Resume CleanUp
End Function

Private Function UpdateMassTotals(cTruck As clsTruck)
    'Update MassANTotal and MassEmulTotal
    Dim dblEmulTot As Double
    Dim dblANTot As Double
    Dim i%
    On Error GoTo errHandler
    
    dblEmulTot = 0#
    dblANTot = 0#
    For i% = 1 To cTruck.Body.Tanks.Count
        If cTruck.Body.Tanks(i%).CurTankUse = ttEmulsion Then
            'Increment Emulsion total
            dblEmulTot = dblEmulTot + cTruck.Body.Tanks(i%).CurVol * cTruck.Body.Tanks(i%).DensityContents
        Else
            'Increment AN total
            dblANTot = dblANTot + cTruck.Body.Tanks(i%).CurVol * cTruck.Body.Tanks(i%).DensityContents
        End If
    Next i%
    cTruck.Body.MassANTotal = dblANTot
    cTruck.Body.MassEmulTotal = dblEmulTot
    Exit Function
errHandler:
    ErrorIn "modSolver.UpdateMassTotals(cTruck)", cTruck
End Function


Private Function FindDistribution(cTruck As clsTruck, Optional IncrFraction As Double = (1# / 100#)) As String
    'Iterates thru product tank-distribution scenarios using supplied
    ' truck configuration.
    'If empty string is returned, a working solution was found and
    ' you can read cTruck to find how it is loaded.  Otherwise, the
    ' return string says why no solution was found.
    'This function has three loops (nested).
    ' - The Emulsion Loop moves emulsion forward.
    ' - The Inner Loop moves AN(2) -> AN(1).
    ' - The Middle Loop moves AN(3) -> [AN(2)+AN(1)].
    Dim sRtn As String
    Dim Loop_e As Double 'Loop counter for Emulsion Loop
    Dim AmtTotal_e As Double 'Total amt to move fm E(2)->E(1)
    Dim Loop_i As Double 'Loop counter for Inner Loop
    Dim AmtTotal_i As Double 'Total amt to move fm AN(2)->AN(1)
    Dim Loop_m As Double 'Loop counter for Middle Loop
    Dim AmtTotal_m As Double 'Total amt to move fm AN(3)-> [AN(2)+AN(1)]
    Dim Loop_o As Double 'Loop counter for Outer Loop
    Dim AmtTotal_o As Double 'Total amt to move fm AN(4)-> [AN(3)+AN(2)+AN(1)]
    
    Dim StepEmul As Double 'amt (volume) of Emul to move on each step through loop
    Dim StepAN As Double 'amt (volume) of AN to move on each step through loop
    Dim dblTemp1 As Double
    Dim dblTemp2 As Double
    Dim cEmulTanks As New clsTanks
    Dim cANTanks As New clsTanks
    
    Dim bFoundASolution As Boolean
    Dim i%
    Dim cLoadingCase As clsLoadingCase 'holds info RE tank loading
    On Error GoTo errHandler
    
    Set cLoadingCase = New clsLoadingCase
    cLoadingCase.InitTanks cTruck
    
    Set cANTanks = ANTanks(cTruck)
    Set cEmulTanks = EmulsionTanks(cTruck)
    
    FindDistribution = "" 'Default to OK
    bFoundASolution = False 'default to FAIL
    
    StepAN = ProductIncrement(cANTanks, IncrFraction) 'AmtTotal_i / 100#
    StepEmul = ProductIncrement(cEmulTanks, IncrFraction) 'AmtTotal_e / 100#
    
    '--- Set bounds for Emulsion Loop (no need to do this more than once)
    If cEmulTanks.Count > 1 Then
        AmtTotal_e = cEmulTanks(1).Volume - cEmulTanks(1).CurVol 'free-space in 1st tank
        If cEmulTanks(2).CurVol < AmtTotal_e Then
            'can't totally fill void in tank 1 since there's not that much emul in tank  2
            AmtTotal_e = cEmulTanks(2).CurVol
        End If
    Else
        'Emulsion Loop unnecessary
        AmtTotal_e = 0#
    End If
    AmtTotal_e = Round(AmtTotal_e, 0) 'round to nearest KG
    
    PrepareOuterLoop cTruck, AmtTotal_m
    'Outer Loop ------------------------------------------------------------------
    For Loop_o = 0# To AmtTotal_o Step StepAN
        'Shift AN(4) Forward
        If (AmtTotal_o <> 0#) And (Loop_o <> 0#) Then
            dblTemp1 = cANTanks(1).CurVol + cANTanks(2).CurVol + cANTanks(3).CurVol + StepAN 'would be put in forward tanks
            dblTemp2 = cANTanks(1).Volume + cANTanks(2).Volume + cANTanks(3).Volume 'capacity of forward tanks
            If (cANTanks(4).CurVol >= StepAN) And (dblTemp2 >= dblTemp1) Then
                'Move AN from this tank (only if applicable)
                cANTanks(4).CurVol = cANTanks(4).CurVol - StepAN
                'Recalculate variables for AN Tank 4
                cANTanks(4).CurStkHt = SolvePolynomial(cANTanks(4).StickLength, cANTanks(4).CurVol)
                cANTanks(4).CurContentCG = SolvePolynomial(cANTanks(4).ContentCG, cANTanks(4).CurStkHt)
                cANTanks(4).CurContentVCG = SolvePolynomial(cANTanks(4).ContentVCG, cANTanks(4).CurStkHt)
            End If
        End If
    
        PrepareMiddleLoop cTruck, AmtTotal_m
        'Middle Loop ------------------------------------------------------------------
        For Loop_m = 0# To AmtTotal_m Step StepAN
            'Shift AN Forward
            If (AmtTotal_m <> 0#) And (Loop_m <> 0#) Then
                dblTemp1 = cANTanks(1).CurVol + cANTanks(2).CurVol + StepAN 'would be put in forward tanks
                dblTemp2 = cANTanks(1).Volume + cANTanks(2).Volume 'capacity of forward tanks
                If (cANTanks(3).CurVol >= StepAN) And (dblTemp2 >= dblTemp1) Then
                    'Move AN from this tank (only if applicable)
                    cANTanks(3).CurVol = cANTanks(3).CurVol - StepAN
                    'Recalculate variables for AN Tank 3
                    cANTanks(3).CurStkHt = SolvePolynomial(cANTanks(3).StickLength, cANTanks(3).CurVol)
                    cANTanks(3).CurContentCG = SolvePolynomial(cANTanks(3).ContentCG, cANTanks(3).CurStkHt)
                    cANTanks(3).CurContentVCG = SolvePolynomial(cANTanks(3).ContentVCG, cANTanks(3).CurStkHt)
                End If
            End If
            
            PrepareInnerLoop cTruck, AmtTotal_i
            'Inner Loop ------------------------------------------------------------------
            For Loop_i = 0# To AmtTotal_i Step StepAN
                'Shift AN Forward
                If (AmtTotal_i <> 0#) And (Loop_i <> 0#) Then
                    dblTemp2 = cANTanks(2).CurVol - StepAN 'remove from 2nd tank
                    dblTemp1 = cANTanks(1).CurVol + StepAN 'move to first tank
                    If (dblTemp2 >= 0) And (dblTemp1 <= cANTanks(1).Volume) Then
                        'Move AN from tank 2 to tank 1
                        cANTanks(2).CurVol = dblTemp2 'cANTanks(2).CurVol - StepAN
                        cANTanks(1).CurVol = dblTemp1 'cANTanks(1).CurVol + StepAN
                        'Recalculate variables for AN Tanks 1&2
                        cANTanks(1).CurStkHt = SolvePolynomial(cANTanks(1).StickLength, cANTanks(1).CurVol)
                        cANTanks(1).CurContentCG = SolvePolynomial(cANTanks(1).ContentCG, cANTanks(1).CurStkHt)
                        cANTanks(1).CurContentVCG = SolvePolynomial(cANTanks(1).ContentVCG, cANTanks(1).CurStkHt)
                        cANTanks(2).CurStkHt = SolvePolynomial(cANTanks(2).StickLength, cANTanks(2).CurVol)
                        cANTanks(2).CurContentCG = SolvePolynomial(cANTanks(2).ContentCG, cANTanks(2).CurStkHt)
                        cANTanks(2).CurContentVCG = SolvePolynomial(cANTanks(2).ContentVCG, cANTanks(2).CurStkHt)
                    End If
                End If
                        
                PrepareEmulsionLoop cTruck
                'Emulsion Loop ------------------------------------------------------------------
                For Loop_e = 0# To AmtTotal_e Step StepEmul
                    'Shift Emulsion Forward
                    If (AmtTotal_e <> 0#) And (Loop_e <> 0#) Then
                        dblTemp2 = cEmulTanks(2).CurVol - StepEmul 'remove from 2nd tank
                        dblTemp1 = cEmulTanks(1).CurVol + StepEmul 'move to first tank
                        If (dblTemp2 >= 0) And (dblTemp1 <= cEmulTanks(1).Volume) Then
                            'Move Emul from tank 2 to tank 1
                            cEmulTanks(2).CurVol = dblTemp2 'cEmulTanks(2).CurVol - StepEmul
                            cEmulTanks(1).CurVol = dblTemp1 'cEmulTanks(1).CurVol + StepEmul
                            'Recalculate variables for Emul Tanks 1&2
                            cEmulTanks(1).CurStkHt = SolvePolynomial(cEmulTanks(1).StickLength, cEmulTanks(1).CurVol)
                            cEmulTanks(1).CurContentCG = SolvePolynomial(cEmulTanks(1).ContentCG, cEmulTanks(1).CurStkHt)
                            cEmulTanks(1).CurContentVCG = SolvePolynomial(cEmulTanks(1).ContentVCG, cEmulTanks(1).CurStkHt)
                        End If
                    End If
                    
                    '=================================================
                    'frmMonitor.MonitorDwg cTruck, sRtn 'Graphical representation of loading
                    '=================================================
    
                    'Recalc Dependant (Component) Tanks
                    sRtn = CalcDependantTanks(cTruck)
                    
                    If sRtn = "" Then
                        'Check Loading since Dependant tanks are OK
                        sRtn = LoadingViolations(cTruck, bIsOnRoad)
                        If sRtn = "" Then
                            'Valid loading!  See if this loading condition is best one yet
                            BestVCOGLoading cTruck, cLoadingCase
                            bFoundASolution = True
                            'GoTo CleanUp
                        End If
                    End If
                    If StepEmul = 0# Then Exit For 'Avoid the zero-step bug
                Next Loop_e ' Emulsion Loop
                If StepAN = 0# Then Exit For 'Avoid the zero-step bug
            Next Loop_i ' Inner Loop
            If StepAN = 0# Then Exit For 'Avoid the zero-step bug
        Next Loop_m ' Middle Loop
        If StepAN = 0# Then Exit For 'Avoid the zero-step bug
    Next Loop_o ' Outer Loop
    
    If bFoundASolution Then
        FindDistribution = ""
        'Set truck to best configuration found
        For i% = 1 To cTruck.Body.Tanks.Count
            cTruck.Body.Tanks(i%).CurVol = cLoadingCase.ContentVol(i%)
        Next i%
        ReCalc cTruck
    Else
        FindDistribution = sRtn
    End If
    
CleanUp:
    Set cANTanks = Nothing
    Set cEmulTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.FindDistribution(cTruck,IncrFraction)", Array(cTruck, IncrFraction)
    Resume CleanUp
End Function


Private Function PrepareEmulsionLoop(cTruck As clsTruck)
    'Called just before the Emulsion Loop that re-distributes Emulsion
    Dim cEmulTanks As New clsTanks
    Dim cANTanks As New clsTanks
    On Error GoTo errHandler

    Set cANTanks = ANTanks(cTruck)
    Set cEmulTanks = EmulsionTanks(cTruck)
            
    'Reset Emulsion so rear tank is most full (if necessary)
    If cEmulTanks.Count > 1 And cANTanks.Count > 1 Then
        'There are 2 Emul tanks and at least one Middle Loop
        If cEmulTanks(2).Volume > (cTruck.Body.MassEmulTotal / cGlobalInfo.DensityEmul) Then
            'Rear Emul tank can hold everything
            cEmulTanks(2).CurVol = cTruck.Body.MassEmulTotal / cGlobalInfo.DensityEmul
            cEmulTanks(1).CurVol = 0#
        Else
            'Fill rear tank then rest goes in front tank
            cEmulTanks(2).CurVol = cEmulTanks(2).Volume
            cEmulTanks(1).CurVol = (cTruck.Body.MassEmulTotal / cGlobalInfo.DensityEmul) - cEmulTanks(2).CurVol
        End If
        'Update tank levels after changing volume
        cEmulTanks(2).CurStkHt = SolvePolynomial(cEmulTanks(2).StickLength, cEmulTanks(2).CurVol)
        cEmulTanks(2).CurContentCG = SolvePolynomial(cEmulTanks(2).ContentCG, cEmulTanks(2).CurStkHt)
        cEmulTanks(2).CurContentVCG = SolvePolynomial(cEmulTanks(2).ContentVCG, cEmulTanks(2).CurStkHt)
        cEmulTanks(1).CurStkHt = SolvePolynomial(cEmulTanks(1).StickLength, cEmulTanks(1).CurVol)
        cEmulTanks(1).CurContentCG = SolvePolynomial(cEmulTanks(1).ContentCG, cEmulTanks(1).CurStkHt)
        cEmulTanks(1).CurContentVCG = SolvePolynomial(cEmulTanks(1).ContentVCG, cEmulTanks(1).CurStkHt)
    End If
CleanUp:
    Set cANTanks = Nothing
    Set cEmulTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.PrepareEmulsionLoop(cTruck)", cTruck
    Resume CleanUp
End Function


Private Function PrepareInnerLoop(cTruck As clsTruck, AmtTotal As Double)
    'Prepares AN tanks 1 & 2
    ' prior to Inner Loop that shifts Tank 2 --> Tank 1
    Dim cANTanks As New clsTanks
    Dim dblVol As Double
    On Error GoTo errHandler

    Set cANTanks = ANTanks(cTruck)
    
    'If there is an Middle Loop, reset AN Tanks 1&2 so tank 2 is most full
    If cANTanks.Count > 2 Then
        If cANTanks.Count = 4 Then
            'Total for Tanks 1&2 = AN_TotalVol - (Tank 3+ Tank 4)
            dblVol = (cTruck.Body.MassANTotal / cGlobalInfo.DensityAN) - (cANTanks(3).CurVol + cANTanks(4).CurVol)
            cANTanks(2).CurVol = dblVol
            cANTanks(1).CurVol = 0#  'TB added 10/29/04 (I believe it was a mistake to leave out)
        Else
            'Total for Tanks 1&2 = AN_TotalVol - Tank 3
            dblVol = (cTruck.Body.MassANTotal / cGlobalInfo.DensityAN) - cANTanks(3).CurVol
            cANTanks(2).CurVol = dblVol
            cANTanks(1).CurVol = 0#  'TB added 10/29/04 (I believe it was a mistake to leave out)
        End If
        If cANTanks(2).CurVol > cANTanks(2).Volume Then
            'Too much for Tank 2 alone, split it up
            cANTanks(2).CurVol = cANTanks(2).Volume
            cANTanks(1).CurVol = dblVol - cANTanks(2).CurVol
        End If
        'Update tank levels after changing volume
        cANTanks(2).CurStkHt = SolvePolynomial(cANTanks(2).StickLength, cANTanks(2).CurVol)
        cANTanks(2).CurContentCG = SolvePolynomial(cANTanks(2).ContentCG, cANTanks(2).CurStkHt)
        cANTanks(2).CurContentVCG = SolvePolynomial(cANTanks(2).ContentVCG, cANTanks(2).CurStkHt)
        cANTanks(1).CurStkHt = SolvePolynomial(cANTanks(1).StickLength, cANTanks(1).CurVol)
        cANTanks(1).CurContentCG = SolvePolynomial(cANTanks(1).ContentCG, cANTanks(1).CurStkHt)
        cANTanks(1).CurContentVCG = SolvePolynomial(cANTanks(1).ContentVCG, cANTanks(1).CurStkHt)
    End If
    
    '--- Set bounds for Inner Loop
    If cANTanks.Count > 1 Then
        AmtTotal = cANTanks(1).Volume - cANTanks(1).CurVol 'free-space in 1st tank
        If cANTanks(2).CurVol < AmtTotal Then
            'can't totally fill void in tank 1 since there's not that much in tank 2
            AmtTotal = cANTanks(2).CurVol
        End If
    Else
        'Inner Loop unnecessary
        AmtTotal = 0#
    End If
    AmtTotal = Round(AmtTotal, 0) 'round to nearest KG
CleanUp:
    Set cANTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.PrepareInnerLoop(cTruck,AmtTotal)", Array(cTruck, AmtTotal)
    Resume CleanUp
End Function


Private Function PrepareMiddleLoop(cTruck As clsTruck, AmtTotal As Double)
    'Prepares AN tanks 1, 2 & 3
    ' prior to Middle Loop that shifts Tank 3 --> (Tank 1 & Tank 2)
    Dim cANTanks As New clsTanks
    Dim dblVol As Double
    On Error GoTo errHandler

    Set cANTanks = ANTanks(cTruck)

    'If there is an Outer Loop, reset AN Tanks 1, 2 & 3 so tank 3 is most full
    If cANTanks.Count = 4 Then
        'Total for Tanks 1,2&3 = AN_TotalVol - Tank 4
        dblVol = (cTruck.Body.MassANTotal / cGlobalInfo.DensityAN) - cANTanks(4).CurVol
        'Shift as much as possible to Tank 3
        If dblVol <= cANTanks(3).Volume Then
            cANTanks(3).CurVol = dblVol
        Else
            cANTanks(3).CurVol = cANTanks(3).Volume
        End If
        'Put remainder in Tank 2
        cANTanks(2).CurVol = dblVol - cANTanks(3).CurVol
        If cANTanks(2).CurVol > cANTanks(2).Volume Then
            'Too much for Tank 2 alone, put remainder in Tank 1
            cANTanks(2).CurVol = cANTanks(2).Volume
            cANTanks(1).CurVol = dblVol - (cANTanks(3).CurVol + cANTanks(2).CurVol)
        End If
        'Update tank levels after changing volume
        cANTanks(3).CurStkHt = SolvePolynomial(cANTanks(3).StickLength, cANTanks(3).CurVol)
        cANTanks(3).CurContentCG = SolvePolynomial(cANTanks(3).ContentCG, cANTanks(3).CurStkHt)
        cANTanks(3).CurContentVCG = SolvePolynomial(cANTanks(3).ContentVCG, cANTanks(3).CurStkHt)
        cANTanks(2).CurStkHt = SolvePolynomial(cANTanks(2).StickLength, cANTanks(2).CurVol)
        cANTanks(2).CurContentCG = SolvePolynomial(cANTanks(2).ContentCG, cANTanks(2).CurStkHt)
        cANTanks(2).CurContentVCG = SolvePolynomial(cANTanks(2).ContentVCG, cANTanks(2).CurStkHt)
        cANTanks(1).CurStkHt = SolvePolynomial(cANTanks(1).StickLength, cANTanks(1).CurVol)
        cANTanks(1).CurContentCG = SolvePolynomial(cANTanks(1).ContentCG, cANTanks(1).CurStkHt)
        cANTanks(1).CurContentVCG = SolvePolynomial(cANTanks(1).ContentVCG, cANTanks(1).CurStkHt)
    End If

    '--- Set bounds for Middle Loop
    If cANTanks.Count > 2 Then
        AmtTotal = cANTanks(1).Volume - cANTanks(1).CurVol _
                   + cANTanks(2).Volume - cANTanks(2).CurVol  'free-space in 1st & 2nd tanks
        If cANTanks(3).CurVol < AmtTotal Then
            'can't totally fill void in tanks 1&2 since there's not that much in tank 3
            AmtTotal = cANTanks(3).CurVol
        End If
    Else
        'Middle Loop unnecessary
        AmtTotal = 0#
    End If
    AmtTotal = Round(AmtTotal, 0) 'round to nearest KG
CleanUp:
    Set cANTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.PrepareMiddleLoop(cTruck,AmtTotal)", Array(cTruck, AmtTotal)
    Resume CleanUp
End Function


Private Function PrepareOuterLoop(cTruck As clsTruck, AmtTotal As Double)
    'Run prior to Outer Loop that shifts Tank 4 --> (Tank 1, Tank 2, & Tank 3)
    ' it figures out how much can be moved from tank 4 forward
    ' (Tank levels are not set because that was already done at top of FindDistribution)
    Dim cANTanks As New clsTanks
    On Error GoTo errHandler

    Set cANTanks = ANTanks(cTruck)

    '--- Set bounds for Outer Loop
    If cANTanks.Count = 4 Then
        AmtTotal = cANTanks(1).Volume - cANTanks(1).CurVol _
                   + cANTanks(2).Volume - cANTanks(2).CurVol _
                   + cANTanks(3).Volume - cANTanks(3).CurVol  'free-space in 1st, 2nd, & 3rd tanks
        
        If cANTanks(4).CurVol < AmtTotal Then
            'can't totally fill void in tanks 1,2,&3 since there's not that much in tank 4
            AmtTotal = cANTanks(4).CurVol
        End If
    Else
        'Outer Loop unnecessary
        AmtTotal = 0#
    End If
    AmtTotal = Round(AmtTotal, 0) 'round to nearest KG
CleanUp:
    Set cANTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.PrepareOuterLoop(cTruck,AmtTotal)", Array(cTruck, AmtTotal)
    Resume CleanUp
End Function


Public Function ReCalc(cTruck As clsTruck) As String
    'Calculate CurStkHt, CurContentCG for each tank.  Also assigns tank levels
    ' to water, gassing, etc.
    ' This func needs to be called right before starting to solve a new
    ' truck config.
    Dim sRtn As String
    Dim dblANTotal As Double
    Dim dblEmulTotal As Double
    Dim i%
    Dim cTank As clsTank
    On Error GoTo errHandler
    
    'Run Calcs for AN & Emul tanks
    For i% = 1 To cTruck.Body.Tanks.Count
        Set cTank = cTruck.Body.Tanks(i%)
        'Calc CurStkHt based on CurVol
        cTank.CurStkHt = SolvePolynomial(cTank.StickLength, cTank.CurVol)
        'Calc CurContentCG based on CurStkHt
        cTank.CurContentCG = SolvePolynomial(cTank.ContentCG, cTank.CurStkHt)
        cTank.CurContentVCG = SolvePolynomial(cTank.ContentVCG, cTank.CurStkHt)
        'Sum product totals
        If cTank.CurTankUse = ttAN Then
            dblANTotal = dblANTotal + cTank.CurVol * cTank.DensityContents
        Else
            dblEmulTotal = dblEmulTotal + cTank.CurVol * cTank.DensityContents
        End If
    Next i%
    'Save updated product totals
    cTruck.Body.MassANTotal = dblANTotal
    cTruck.Body.MassEmulTotal = dblEmulTotal
    'Set levels for dependent tanks
    sRtn = CalcDependantTanks(cTruck)
    
    ReCalc = sRtn 'may return error(s) indicating that dependant tanks not big enough
    Set cTank = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.ReCalc(cTruck)", cTruck
End Function


Private Function CalcDependantTanks(cTruck As clsTruck) As String
    'sets the dependant tank (water, gassing, etc.) volumes, levels, and CGs
    Dim i%
    Dim x%
    Dim sRtn As String
    Dim dblExcess As Double
    Dim sExcess As String
    Dim dblMass As Double
    Dim cComponent As clsComponent
    On Error GoTo errHandler
    
    CalcDependantTanks = "" 'default to OK
    For i% = 1 To cTruck.Components.Count
        Set cComponent = cTruck.Components(i%)
        If cComponent.FillRelationShips.Count > 0 Then
            Select Case FillMethodThisComponent(cComponent.ContentsType)
            Case 0 'Auto
                'Set fill-level based on defined relationships
                dblMass = 0#
                For x% = 1 To cComponent.FillRelationShips.Count
                    Select Case UCase$(cComponent.FillRelationShips(x%).ParentProduct)
                    Case ptAN
                        dblMass = dblMass + cComponent.FillRelationShips(x%).Multiplier _
                                   * cTruck.Body.MassANTotal + cComponent.FillRelationShips(x%).Offset
                    Case ptEmulsion
                        dblMass = dblMass + cComponent.FillRelationShips(x%).Multiplier _
                                   * cTruck.Body.MassEmulTotal + cComponent.FillRelationShips(x%).Offset
                    End Select
                Next x%
                'Convert the mass into volume
                cComponent.Capacity.CurVol = dblMass / cComponent.Capacity.DensityContents
                If cComponent.Capacity.CurVol > cComponent.Capacity.Volume Then
                    'Tank was over-full.  Set error and Normalize.
                    dblExcess = (cComponent.Capacity.CurVol - cComponent.Capacity.Volume) _
                                * cGlobalInfo.VolumeUnits.Multiplier
                    sExcess = Format(dblExcess, "### ") & cGlobalInfo.VolumeUnits.Display
                    sRtn = sRtn & "Capacity of " & cComponent.DisplayName & _
                           " short by " & sExcess & vbCrLf
                    cComponent.Capacity.CurVol = cComponent.Capacity.Volume
                End If
            Case 1 'Empty
                cComponent.Capacity.CurVol = 0#
            Case 2 'Full
                cComponent.Capacity.CurVol = cComponent.Capacity.Volume
            End Select
            
            'Calculate level and CG
            cComponent.Capacity.CurStkHt = SolvePolynomial( _
                                            cComponent.Capacity.StickLength _
                                            , cComponent.Capacity.CurVol)
            cComponent.Capacity.CurContentCG = SolvePolynomial( _
                                            cComponent.Capacity.ContentCG _
                                            , cComponent.Capacity.CurStkHt) _
                                            + cComponent.Offset
        ElseIf Var2Dbl(cComponent.Capacity.DefaultVolContents) > 0 Then
            'No fill Relationships & calculate using Default(user) content Volume
            cComponent.Capacity.CurContentCG = cComponent.EmptyCG + cComponent.Offset
        ElseIf Var2Dbl(cComponent.Capacity.DefaultWtContents) > 0 Then
            'No fill Relationships & calculate using Default(user) content Weight
            cComponent.Capacity.CurContentCG = cComponent.EmptyCG + cComponent.Offset
        End If
    Next i%
    
    Set cComponent = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.CalcDependantTanks(cTruck)", cTruck
End Function


Private Function FillMethodThisComponent(ContentsType As ComponentContentsType) As Integer
    Select Case ContentsType
    Case ctFuel
        FillMethodThisComponent = cGlobalInfo.FillMethod_Fuel
    Case ctWater
        FillMethodThisComponent = cGlobalInfo.FillMethod_Water
    Case ctGasA
        FillMethodThisComponent = cGlobalInfo.FillMethod_GasA
    Case ctGasB
        FillMethodThisComponent = cGlobalInfo.FillMethod_GasB
    Case ctAdditive
        FillMethodThisComponent = cGlobalInfo.FillMethod_Additive
    End Select
End Function


Public Function EmulsionTanks(cTruck As clsTruck) As clsTanks
    'Returns a (linked) collection of the Emulsion Tanks
    Dim i%
    Dim cTank As clsTank
    Dim cEmTanks As clsTanks
    On Error GoTo errHandler
    
    Set cEmTanks = New clsTanks
    For i% = 1 To cTruck.Body.Tanks.Count
        Set cTank = cTruck.Body.Tanks(i%)
        If cTank.CurTankUse = ttEmulsion Then
            cEmTanks.Add cTank
        End If
    Next i%
    Set EmulsionTanks = cEmTanks
CleanUp:
    Set cTank = Nothing
    Set cEmTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.EmulsionTanks(cTruck)", cTruck
    Resume CleanUp
End Function

Public Function ANTanks(cTruck As clsTruck) As clsTanks
    'Returns a (linked) collection of the Emulsion Tanks
    Dim i%
    Dim cTank As clsTank
    Dim cANTanks As clsTanks
    On Error GoTo errHandler
    
    Set cANTanks = New clsTanks
    For i% = 1 To cTruck.Body.Tanks.Count
        Set cTank = cTruck.Body.Tanks(i%)
        If cTank.CurTankUse = ttAN Then
            cANTanks.Add cTank
        End If
    Next i%
    Set ANTanks = cANTanks
CleanUp:
    Set cTank = Nothing
    Set cANTanks = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.ANTanks(cTruck)", cTruck
    Resume CleanUp
End Function

Public Function SetComponentDensities(cTruck As clsTruck) As String
    'Call this function to set density property of components based on GlobalInfo
    Dim i%
    Dim lmt As Integer
    On Error GoTo errHandler
    
    lmt = cTruck.Components.Count
    For i% = 1 To lmt
        Select Case cTruck.Components(i%).ContentsType
        Case ctNone
            'Do nothing
        Case ctFuel
            cTruck.Components(i%).Capacity.DensityContents = cGlobalInfo.DensityFuel
        Case ctWater
            cTruck.Components(i%).Capacity.DensityContents = cGlobalInfo.DensityWater
        Case ctGasA
            cTruck.Components(i%).Capacity.DensityContents = cGlobalInfo.DensityGasA
        Case ctGasB
            cTruck.Components(i%).Capacity.DensityContents = cGlobalInfo.DensityGasB
        Case ctAdditive
            cTruck.Components(i%).Capacity.DensityContents = cGlobalInfo.DensityAdditive
        Case ctOther
            'Do nothing
        End Select
    Next
    SetComponentDensities = ""
    Exit Function
errHandler:
    SetComponentDensities = Err.Description
End Function


Public Function SetProductDensities(cTruck As clsTruck) As String
    'Call this function to set density property of components based on GlobalInfo
    Dim i%
    Dim lmt As Integer
    On Error GoTo errHandler

    lmt = cTruck.Body.Tanks.Count
    For i% = 1 To lmt
        Select Case cTruck.Body.Tanks(i%).CurTankUse
        Case ttAN
            cTruck.Body.Tanks(i%).DensityContents = cGlobalInfo.DensityAN
        Case Else
            cTruck.Body.Tanks(i%).DensityContents = cGlobalInfo.DensityEmul
        End Select
    Next
    SetProductDensities = ""
    Exit Function
errHandler:
    SetProductDensities = Err.Description
End Function

Private Function ProductIncrement(cTanks As clsTanks, Optional IncrFraction As Double = (1# / 200#)) As Double
    'Returns a (volum) number that's approximately equal to 1/200 (or specified fraction)
    ' the volume of the smallest tank
    Dim i%
    Dim dblTemp As Double
    Dim dblRtn As Double
    On Error GoTo errHandler
    
    If cTanks.Count = 0 Then
        ProductIncrement = 0#
        Exit Function
    End If
    
    dblRtn = 9E+99 'start with LARGE value
    For i% = 1 To cTanks.Count
        dblTemp = cTanks(i%).Volume * IncrFraction
        If dblTemp < dblRtn Then
            dblRtn = dblTemp
        End If
    Next i%
    ProductIncrement = dblRtn
    Exit Function
errHandler:
    ErrorIn "modSolver.ProductIncrement(cTanks,IncrFraction)", Array(cTanks, IncrFraction)
End Function

Public Function CurrentConfigVal(cTruck As clsTruck) As Integer
    'return current configuration for provided truck
    Dim NumDualUse As Integer
    Dim intConfig  As Integer
    Dim i%
    On Error GoTo errHandler
    
    NumDualUse = 0
    intConfig = 0
    For i% = 1 To cTruck.Body.Tanks.Count
        If cTruck.Body.Tanks(i%).TankType = ttDual Then
            intConfig = intConfig + _
                   cTruck.Body.Tanks(i%).CurTankUse * ((2 ^ NumDualUse))
            NumDualUse = NumDualUse + 1
        End If
    Next i%
    CurrentConfigVal = intConfig
    Exit Function
errHandler:
    ErrorIn "modSolver.CurrentConfigVal(cTruck)", cTruck
End Function

Private Function AdaptiveEmulPct(cTruck As clsTruck) As Double
    'Although the user may have specified a value for the global
    ' variable 'dblEmulPct', there are certain calculations that
    ' will fail if an adjustment isn't made for situations where
    ' 'dblEmulPct' is unreasonable.  Even though the user may
    ' specify 30%, that won't work if all tanks are setup as AN.
    Dim i%
    Dim intTemp As Integer
    On Error GoTo errHandler
    
    'Change target Emul% for this truck  if necessary
    For i% = 1 To cTruck.Body.Tanks.Count
        intTemp = intTemp + cTruck.Body.Tanks(i%).CurTankUse
    Next i%
    If intTemp = 0 Then
        'All AN tanks, reset Emul% to zero
        AdaptiveEmulPct = 0#
    ElseIf intTemp = cTruck.Body.Tanks.Count Then
        'All Emulsion tanks, reset Emul% to 100
        AdaptiveEmulPct = 100#
    Else
        'Use given Emul%
        AdaptiveEmulPct = dblEmulPct
    End If
    Exit Function
errHandler:
    ErrorIn "modSolver.AdaptiveEmulPct(cTruck)", cTruck
End Function

Public Function SolvePolynomial(colK As Collection, dblInput) As Double
    Dim k5 As Double
    Dim k4 As Double
    Dim k3 As Double
    Dim k2 As Double
    Dim k1 As Double
    Dim k0 As Double
    Dim x As Double
    Dim x2 As Double
    Dim x3 As Double
    Dim x4 As Double
    Dim x5 As Double
    On Error GoTo errHandler
    
    x = dblInput
    k5 = CDbl(colK.Item("5"))
    k4 = CDbl(colK.Item("4"))
    k3 = CDbl(colK.Item("3"))
    k2 = CDbl(colK.Item("2"))
    k1 = CDbl(colK.Item("1"))
    k0 = CDbl(colK.Item("0"))
    
    'Too slow...
    'SolvePolynomial = k5 * x ^ 5 + k4 * x ^ 4 + k3 * x ^ 3 + k2 * x ^ 2 + k1 * x + k0

    x2 = x ^ 2
    x3 = x2 * x
    x4 = x3 * x
    x5 = x4 * x
    SolvePolynomial = k5 * x5 + k4 * x4 + k3 * x3 + k2 * x2 + k1 * x + k0
    Exit Function
errHandler:
    ErrorIn "modSolver.SolvePolynomial(colK,dblInput)", Array(colK, dblInput)
End Function

Private Sub BestVCOGLoading(cTruck As clsTruck, cLoadingCase As clsLoadingCase)
    'This routine saves the current config data if the loading is best one yet
    Dim i%
    Dim dblVCOG As Double 'moment of current config
    Dim cTank As clsTank
    On Error GoTo errHandler

    dblVCOG = ProductVCOG(cTruck.Body)
    
    If dblVCOG > cLoadingCase.VCOG Then Exit Sub 'No better. Exit.
    
    'This is a better loading config, so save it in class
    For i% = 1 To cTruck.Body.Tanks.Count
        Set cTank = cTruck.Body.Tanks(i%)
        cLoadingCase.SetTankConfig cTank.CurTankUse, CStr(i%)
        cLoadingCase.SetContentVol cTank.CurVol, CStr(i%)
    Next i%
    cLoadingCase.VCOG = dblVCOG
    Set cTank = Nothing
    Exit Sub
errHandler:
    ErrorIn "modSolver.BestVCOGLoading(cTruck,cLoadingCase)", Array(cTruck, cLoadingCase)
End Sub

Private Sub HighestCapacityLoading(cTruck As clsTruck, cLoadingCase As clsLoadingCase)
    'This routine saves the current config data if the loading is best one yet
    Dim i%
    Dim dblMass As Double
    Dim cTank As clsTank
    On Error GoTo errHandler

    dblMass = cTruck.Body.MassANTotal + cTruck.Body.MassEmulTotal
    
    If dblMass < cLoadingCase.TotalProductMass Then Exit Sub 'No better. Exit.
    
    'This is a better loading config, so save it in class
    For i% = 1 To cTruck.Body.Tanks.Count
        Set cTank = cTruck.Body.Tanks(i%)
        cLoadingCase.SetTankConfig cTank.CurTankUse, CStr(i%)
        cLoadingCase.SetContentVol cTank.CurVol, CStr(i%)
    Next i%
    cLoadingCase.TotalProductMass = dblMass
    Set cTank = Nothing
    Exit Sub
errHandler:
    ErrorIn "modSolver.HighestCapacityLoading(cTruck,cLoadingCase)", Array(cTruck, cLoadingCase)
End Sub


Public Function ProductVCOG(cBody As clsBody) As Double 'meters
    Dim i%
    Dim dblMass As Double
    Dim dblTotalMass As Double
    Dim dblMoment As Double 'moment of current config
    Dim cTank As clsTank
    On Error GoTo errHandler
    
    dblMoment = 0#
    dblTotalMass = 0#
    For i% = 1 To cBody.Tanks.Count
        Set cTank = cBody.Tanks(i%)
        dblMass = cTank.CurVol * cTank.DensityContents
        dblMoment = dblMoment + (dblMass * cTank.CurContentVCG)
        dblTotalMass = dblTotalMass + dblMass
    Next i%

    If dblTotalMass > 0 Then
        ProductVCOG = dblMoment / dblTotalMass
    Else
        'Avoid division by zero
        ProductVCOG = 0
    End If
    Set cTank = Nothing
    Exit Function
errHandler:
    ErrorIn "modSolver.ProductVCOG(cBody)", cBody
End Function


Public Function IsBlendConfiguration(cTruck As clsTruck) As Boolean
    'Returns TRUE (only) if it makes sense to display emulsion percent
    'A FALSE will be returned if cTruck is configured for all ANFO or all Emul
    Dim i%
    Dim intVal
    On Error GoTo errHandler
    
    For i% = 1 To cTruck.Body.Tanks.Count
        intVal = intVal + cTruck.Body.Tanks(i%).CurTankUse
    Next i%
    If intVal = 0 Or intVal = cTruck.Body.Tanks.Count Then
        IsBlendConfiguration = False
    Else
        IsBlendConfiguration = True
    End If
    Exit Function
errHandler:
    ErrorIn "modSolver.IsBlendConfiguration(cTruck)", cTruck
End Function

