Attribute VB_Name = "modDXF"
Option Explicit

Public cDXF As DXFReaderClass


Public Sub RenderTruck(cTruck As clsTruck, ctlDXF As DXFReader, bCurbSide As Boolean)
    Dim sPath As String
    Dim sFile As String
    Dim xLoc As Single
    Dim yLoc As Single
    Dim xEnd As Single
    Dim TagDia As Single
    Dim intC As Integer
    Dim i%
    
    sPath = AddBackslash(cGlobalInfo.ALObjectFolder) & "DXFs\"
    
    Set cDXF = New DXFReaderClass
    With ctlDXF
        .Clear
        .NewDrawing
        '.ReadDXF sPath & "Blank.dxf"
        
        'Draw Chassis
        If bCurbSide Then
            sFile = cTruck.Chassis.CurbSideStd
        Else
            sFile = cTruck.Chassis.StreetSideStd
        End If
        xLoc = 0
        intC = 1
        If sFile <> "" Then
            ReadComponent sPath & sFile, intC, xLoc, ctlDXF
        Else
            MsgBox "Chassis component has no associated drawing.", vbInformation, "Notice"
        End If
        
        'Draw Wheels
        If cTruck.Chassis.TandemSpacing = 0 Then
            'Single rear axle
            xLoc = (cTruck.Chassis.WB * 39.37)
            If bCurbSide Then xLoc = -xLoc
            cDXF.AddCircle ctlDXF, xLoc, cTruck.Chassis.WheelY * 39.37, 0, cTruck.Chassis.WheelDia / 2 * 39.37
        Else
            'Tandem
            xLoc = (cTruck.Chassis.WB - cTruck.Chassis.TandemSpacing / 2) * 39.37
            If bCurbSide Then xLoc = -xLoc
            cDXF.AddCircle ctlDXF, xLoc, cTruck.Chassis.WheelY * 39.37, 0, cTruck.Chassis.WheelDia / 2 * 39.37
            xLoc = (cTruck.Chassis.WB + cTruck.Chassis.TandemSpacing / 2) * 39.37
            If bCurbSide Then xLoc = -xLoc
            cDXF.AddCircle ctlDXF, xLoc, cTruck.Chassis.WheelY * 39.37, 0, cTruck.Chassis.WheelDia / 2 * 39.37
        End If
        
        'Draw Tag(s)
        For i% = 1 To cTruck.Chassis.Tags.Count
            xLoc = cTruck.Chassis.Tags(i%).Location * 39.37
            If bCurbSide Then xLoc = -xLoc
            TagDia = 0.8 * cTruck.Chassis.WheelDia
            If cTruck.Chassis.Tags(i%).DownwardForce > 0 Then
                'Tag is down
                yLoc = cTruck.Chassis.WheelY - (cTruck.Chassis.WheelDia - TagDia) / 2
            Else
                'Tag is up
                yLoc = -TagDia / 2
            End If
            cDXF.AddCircle ctlDXF, xLoc, yLoc * 39.37, 0, TagDia * 39.37 / 2
        Next i%
        
        '.Regen
        '.ZoomExtents
        DoEvents
        
       'Draw Body
        If bCurbSide And (cTruck.Body.Placement <> plStreetSideStd) Then
            'Show a curb-side view
            sFile = cTruck.Body.CurbSideStd
            xLoc = -cTruck.BodyLocation * 39.37
        ElseIf (cTruck.Body.Placement <> plCurbSideStd) Then
            'Show a street-side view
            sFile = cTruck.Body.StreetSideStd
            xLoc = cTruck.BodyLocation * 39.37
        End If
        intC = 2
        If sFile <> "" Then ReadComponent sPath & sFile, intC, xLoc, ctlDXF
        '.Regen
        DoEvents
        
        'Draw Components
        For i% = 1 To cTruck.Components.Count
            If bCurbSide And (cTruck.Components(i%).Placement <> plStreetSideStd) _
            And (cTruck.Components(i%).Placement <> plNotPlaced) Then
                'Show a curb-side view
                sFile = cTruck.Components(i%).CurbSideStd
                xLoc = -cTruck.Components(i%).Offset * 39.37
                If cTruck.Components(i%).LocationReference = orBodyOrigin Then
                    xLoc = xLoc - cTruck.BodyLocation * 39.37
                End If
            ElseIf (Not bCurbSide) And (cTruck.Components(i%).Placement <> plCurbSideStd) _
            And (cTruck.Components(i%).Placement <> plNotPlaced) Then
                'Show a street-side view
                sFile = cTruck.Components(i%).StreetSideStd
                xLoc = cTruck.Components(i%).Offset * 39.37
                If cTruck.Components(i%).LocationReference = orBodyOrigin Then
                    xLoc = xLoc + cTruck.BodyLocation * 39.37
                End If
            End If
            intC = intC + 1
            If sFile <> "" Then ReadComponent sPath & sFile, intC, xLoc, ctlDXF
            '.Regen
            DoEvents
        Next i%
        .ZoomExtents
    
        'Draw Frame Rail (4 inches short of extreme X)
        If bCurbSide Then
            xLoc = -cTruck.BodyLocation * 39.37
            xEnd = .ViewPortMinX + 4
            cDXF.AddLine ctlDXF, xLoc, 0, 0, xEnd, 0, 0
            cDXF.AddLine ctlDXF, xEnd, 0, 0, _
                                 xEnd, -cTruck.Chassis.FrameRailHt * 39.37, 0
            cDXF.AddLine ctlDXF, xEnd, -cTruck.Chassis.FrameRailHt * 39.37, 0, _
                                 xLoc, -cTruck.Chassis.FrameRailHt * 39.37, 0
        Else
            xLoc = cTruck.BodyLocation * 39.37
            xEnd = .ViewPortMaxX - 4
            cDXF.AddLine ctlDXF, xLoc, 0, 0, xEnd, 0, 0
            cDXF.AddLine ctlDXF, xEnd, 0, 0, _
                                 xEnd, -cTruck.Chassis.FrameRailHt * 39.37, 0
            cDXF.AddLine ctlDXF, xEnd, -cTruck.Chassis.FrameRailHt * 39.37, 0, _
                                 xLoc, -cTruck.Chassis.FrameRailHt * 39.37, 0
        End If
        '.Regen
        .ZoomExtents
    End With
End Sub


Private Sub ReadComponent(sDXFFile As String, intNewPrefix As Integer, xLoc As Single, ctlDXF As DXFReader)
    Dim Block As DXFReaderBlock
    Dim Entity As DXFReaderEntity
    Dim n$
    
    With ctlDXF
        .ReadDXFEntities sDXFFile, , , , xLoc
        'Rename blocks to be unique to this DXF file
'        For Each Block In .Blocks
'            n$ = Block.BlockName
'            If Mid$(n$, 1, 1) <> "Z" Then
'                Block.BlockName = "Z" & CStr(intNewPrefix) & n$
'                ChangeKey Block, "Z" & CStr(intNewPrefix) & n$, .Blocks
'            End If
'        Next Block
'        For Each Entity In .entities
'            If Entity.EntityType = "INSERT" Then
'                n$ = Entity.BlockName
'                If Mid$(n$, 1, 1) <> "Z" Then
'                    Entity.BlockName = "Z" & CStr(intNewPrefix) & n$
'                End If
'            End If
'        Next Entity
    End With
End Sub

Private Function ChangeKey(Object As Object, NewKey As String, Collection As Collection) As Boolean
    Dim Item As Object
    Dim Index As Long
    
    For Each Item In Collection
        Index = Index + 1
        If Item Is Object Then
            If Len(NewKey) Then
                Collection.Add Object, NewKey, , Index
            Else
                Collection.Add Object, , , Index
            End If
            Collection.Remove Index
            ChangeKey = True
            Exit For
        End If
    Next
End Function



