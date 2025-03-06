Attribute VB_Name = "modALObjects"
Option Explicit
'This module contains routines to deal with ALObjects -- XML chunks that
' represent parts of a truck.
' At this point the following ALObject types are defined by their
' XML "ObjectType" attribute:
'   - Body
'   - Chassis
'   - Component

Public Function RtnALObject(strFile As String) As clsALObject
    'Read a file and return the high-level object info
    Dim cALObject As clsALObject
    Dim xDoc As MSXML2.DOMDocument30
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xAttr As IXMLDOMAttribute
    On Error GoTo errHandler
    
    Set xDoc = New MSXML2.DOMDocument30
    xDoc.async = False 'Don't load asynchronysly
    If xDoc.Load(strFile) Then
       ' The document loaded successfully.
        Set cALObject = New clsALObject
        'Get ObjectType node
        Set xNode = xDoc.selectSingleNode("ALObject")
        If TypeName(xNode) = "Nothing" Then
            Set RtnALObject = Nothing
            GoTo CleanUp
        End If
        'Read ObjectType attribute
        Set xAttr = xNode.Attributes.getNamedItem("ObjectType")
        If TypeName(xAttr) = "Nothing" Then
            Set RtnALObject = Nothing
            GoTo CleanUp
        End If
        Select Case UCase$(xAttr.nodeValue)
        Case "BODY"
            cALObject.ObjectType = xptBody
        Case "CHASSIS"
            cALObject.ObjectType = xptChassis
        Case "COMPONENT"
            cALObject.ObjectType = xptComponent
        End Select
        'Read FullName node
        Set xAttr = xNode.Attributes.getNamedItem("FullName")
        If TypeName(xAttr) = "Nothing" Then
            Set RtnALObject = Nothing
            GoTo CleanUp
        End If
        cALObject.FullName = xAttr.nodeValue
        'Read DisplayName node (or use truncated filename) for ALObject.DisplayName property
        Set xAttr = xNode.Attributes.getNamedItem("DisplayName")
        If TypeName(xAttr) <> "Nothing" Then
            cALObject.DisplayName = xAttr.nodeValue
        Else
            'No Displayname, so set as displayname=filename
            cALObject.DisplayName = GetFileBaseName(strFile)
            If InStr(cALObject.DisplayName, "_") > 0 Then
                'Strip off front
                cALObject.DisplayName = Mid$(cALObject.DisplayName, InStr(cALObject.DisplayName, "_") + 1)
            End If
        End If
        
        cALObject.File = strFile
        Set RtnALObject = cALObject
    Else
       ' The document failed to load.
        Set RtnALObject = Nothing
    End If

CleanUp:
    Set cALObject = Nothing
    Set xNode = Nothing
    Set xDoc = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.RtnALObject(strFile)", strFile
    Resume CleanUp
End Function



Public Function ALObjectsInPath(strPath As String, ObjectType As ALObjectType) As clsALObjects
    Dim fso As New Scripting.FileSystemObject
    Dim fld As Scripting.Folder
    Dim fl As Scripting.File
    Dim cALObject As clsALObject
    Dim cALObjects As New clsALObjects
    Dim lIndex As Long
    On Error GoTo errHandler
    
    Set fld = fso.GetFolder(strPath)
    lIndex = 0
    For Each fl In fld.Files
        If UCase$(Right$(fl.Name, 4)) = ".XML" Then
            Set cALObject = New clsALObject
            Set cALObject = RtnALObject(fl.Path)
            If TypeName(cALObject) <> "Nothing" Then
                'Valid object, now see if it's the type we're looking for
                If cALObject.ObjectType = ObjectType Then
                    cALObject.Index = lIndex
                    cALObjects.Add cALObject, CStr(cALObject.Index)
                    lIndex = lIndex + 1
                End If
            End If
        End If
    Next
    Set ALObjectsInPath = cALObjects
    
CleanUp:
    Set fso = Nothing
    Set fld = Nothing
    Set fl = Nothing
    Set cALObject = Nothing
    Set cALObjects = Nothing
    Exit Function
errHandler:
    ErrorIn "modALObjects.ALObjectsInPath(strPath,ObjectType)", Array(strPath, ObjectType)
    Resume CleanUp
End Function

Public Function ReplaceChassis(cTruck As clsTruck, sFile As String) As Boolean
    'Replaces the chassis in cTruck with the chassis info from sFile
    Dim xDoc As MSXML2.DOMDocument30
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim sMsg As String
    Dim sErrMsg As String
    Dim sExistingChassisGUID As String
    On Error GoTo errHandler
    
    ReplaceChassis = False 'Default to FAIL
    Set xDoc = New MSXML2.DOMDocument30
    xDoc.async = False 'Don't load asynchronysly
    
    If xDoc.Load(sFile) Then
       ' The document loaded successfully.
        Set xNode = xDoc.selectSingleNode("ALObject")
        If TypeName(xNode) = "Nothing" Then
            'Not an AxleLoad object file!
            GoTo CleanUp
        End If
        Set xNode = xNode.selectSingleNode("Chassis")
        If TypeName(xNode) = "Nothing" Then
            'Couldn't find Chassis node
            GoTo CleanUp
        End If
        'Replace the existing Chassis (and tags)
        sExistingChassisGUID = cTruck.Chassis.GUID
        sErrMsg = ParseChassis(xNode.childNodes, cTruck)
        If sErrMsg <> "" Then
            MsgBox "Corrupt Chassis File.  Please restart program.", vbCritical
            GoTo CleanUp
        End If
    
        'Does ALObject file contain additional Components? (e.g. fuel, driver, etc.)
        Set xNode = xDoc.selectSingleNode("ALObject")
        Set xNode = xNode.selectSingleNode("Components")
        If TypeName(xNode) <> "Nothing" Then
            'See if user wants to import Components
            sMsg = frmChassisComponent.ImportOption
            Select Case sMsg
            Case "IGNORE"
                'do nothing
            Case "ADD"
                ParseComponents xNode.childNodes, cTruck, True
            Case "REPLACE"
                RemoveChassisComponents cTruck, sExistingChassisGUID
                'Now load new chassis-related components
                ParseComponents xNode.childNodes, cTruck, True
            End Select
        End If
    Else
       ' The document failed to load.
        GoTo CleanUp
    End If
    ReplaceChassis = True 'OK if we got this far
    
CleanUp:
    Set xNode = Nothing
    Set xNodes = Nothing
    Set xDoc = Nothing
    Exit Function
errHandler:
    ErrorIn "modALObjects.ReplaceChassis(cTruck,sFile)", Array(cTruck, sFile)
    Resume CleanUp
End Function


Public Function ReplaceBody(cTruck As clsTruck, sFile As String) As Boolean
    'Replaces the body in cTruck with the body info from sFile
    Dim xDoc As MSXML2.DOMDocument30
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim sMsg As String
    Dim sErrMsg As String
    On Error GoTo errHandler
    
    ReplaceBody = False 'Default to FAIL
    Set xDoc = New MSXML2.DOMDocument30
    xDoc.async = False 'Don't load asynchronysly
    
    If xDoc.Load(sFile) Then
       ' The document loaded successfully.
        Set xNode = xDoc.selectSingleNode("ALObject")
        If TypeName(xNode) = "Nothing" Then
            'Not an AxleLoad object file!
            GoTo CleanUp
        End If
        Set xNode = xNode.selectSingleNode("Body")
        If TypeName(xNode) = "Nothing" Then
            'Couldn't find Chassis node
            GoTo CleanUp
        End If
        'Replace the existing Chassis (and tags)
        sErrMsg = ParseBody(xNode.childNodes, cTruck)
        If sErrMsg <> "" Then
            MsgBox "Corrupt Body File.  Please restart program.", vbCritical
            GoTo CleanUp
        End If
    
        'Does ALObject file contain additional Components? (e.g. fuel, driver, etc.)
        Set xNode = xDoc.selectSingleNode("ALObject")
        Set xNode = xNode.selectSingleNode("Components")
        If TypeName(xNode) <> "Nothing" Then
            'See if user wants to import Components
            sMsg = "The selected Body File also contains additional Component data" & vbCrLf & vbCrLf & _
                   "Do you want to import this data with the selected chassis?"
            If MsgBox(sMsg, vbYesNo, "Import body-related components?") = vbYes Then
                '*** Import Components here
                ParseComponents xNode.childNodes, cTruck
            End If
        End If
    Else
       ' The document failed to load.
        GoTo CleanUp
    End If
    ReplaceBody = True 'OK if we got this far
    
CleanUp:
    Set xNode = Nothing
    Set xNodes = Nothing
    Set xDoc = Nothing
    Exit Function
errHandler:
    ErrorIn "modALObjects.ReplaceBody(cTruck,sFile)", Array(cTruck, sFile)
    Resume CleanUp
End Function

Public Function LoadComponent(cTruck As clsTruck, sFile As String) As Boolean
    'Loads the Component object from sFile into cTruck
    Dim xDoc As MSXML2.DOMDocument30
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim sErrMsg As String
    On Error GoTo errHandler
    
    LoadComponent = False 'Default to FAIL
    Set xDoc = New MSXML2.DOMDocument30
    xDoc.async = False 'Don't load asynchronysly
    
    If xDoc.Load(sFile) Then
       ' The document loaded successfully.
        Set xNode = xDoc.selectSingleNode("ALObject")
        If TypeName(xNode) = "Nothing" Then
            'Not an AxleLoad object file!
            GoTo CleanUp
        End If
        Set xNode = xNode.selectSingleNode("Component")
        If TypeName(xNode) = "Nothing" Then
            'Couldn't find Component node
            GoTo CleanUp
        End If
        'add the Component object to the Truck object
        sErrMsg = ParseComponent(xNode, cTruck)
        If sErrMsg <> "" Then
            MsgBox "Corrupt Component File.  Please restart program.", vbCritical
            GoTo CleanUp
        End If
    
        LoadComponent = True 'OK if we got this far
    Else
        LoadComponent = False 'Didn't succeed at loading the file
    End If
CleanUp:
    Set xNode = Nothing
    Set xNodes = Nothing
    Set xDoc = Nothing
    Exit Function
    
errHandler:
    ErrorIn "modALObjects.LoadComponent(cTruck,sFile)", Array(cTruck, sFile)
    Resume CleanUp
End Function


Public Function SaveChassisObject(cTruck As clsTruck, colComponents As Collection, DisplayName As String, FullName As String) As String
    'Saves the Chassis info of given truck to an object file
    Dim xDoc As MSXML2.DOMDocument30
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim xAttribute As MSXML2.IXMLDOMAttribute
    Dim sFile As String
    Dim sMsg As String
    Dim sErrMsg As String
    Dim i%
    Dim sNodeName() As String
    On Error GoTo errHandler
    
    SaveChassisObject = "Unknown Error in modALObjects:SaveChassisObject" 'default to fail
    
    'Create a default file name
    If Trim$(DisplayName) <> "" Then
        sFile = "Chassis_" & Trim$(DisplayName) & ".XML"
    Else
        sFile = "Chassis_" & cTruck.Chassis.DisplayName & ".XML"
    End If
    sFile = FixedFileName(sFile)
    
    With frmMain.dlgFile
        .CancelError = True
        .InitDir = cGlobalInfo.ALObjectFolder
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        .FileName = sFile
        .Filter = "Chassis Files|Chassis_*.xml"
        On Error Resume Next
        .ShowSave
        'If User selected a file, open it
        If Err <> 0 Then
            'User canceled
            SaveChassisObject = "File save was aborted."
            On Error GoTo 0
            GoTo CleanUp
        End If
        On Error GoTo errHandler
        sFile = Trim$(.FileName)
        
        If ValidObjectFileName("Chassis_", sFile) <> "" Then
            'bad file name
            MsgBox ValidObjectFileName("Chassis_", sFile), vbCritical, "Error"
            'recursive call to give user 2nd chance
            SaveChassisObject = SaveChassisObject(cTruck, colComponents, DisplayName, FullName)
        ElseIf FileExists(sFile) Then
            'See if user wants to over-write existing file
            sMsg = "The chosen filename exists.  " & vbCrLf & _
                   "Choose 'Yes' to overwrite existing file," & vbCrLf & _
                   "or 'No' to choose another name"
            If MsgBox(sMsg, vbYesNo, "Overwrite existing file?") = vbNo Then
                'recursive call to give user 2nd chance
                SaveChassisObject = SaveChassisObject(cTruck, colComponents, DisplayName, FullName)
                Exit Function
            End If
        End If
    End With
    On Error GoTo errHandler
    
    'Now save to indicated file ============================================
    'Create new DOM Document
    Set xDoc = New MSXML2.DOMDocument30
    With xDoc
        .async = False
        .validateOnParse = False
        .resolveExternals = False
        .preserveWhiteSpace = True
    End With
    'create processing instructions
    Set xNode = xDoc.createProcessingInstruction("xml", "version='1.0'")
    xDoc.appendChild xNode
    
    'Create Base Node & attributes
    Set xParent = xDoc.createElement("ALObject")
    ' create ObjectType attribute
    Set xAttribute = xDoc.createAttribute("ObjectType")
    xAttribute.Text = "Chassis"
    xParent.Attributes.setNamedItem xAttribute
    
    ' create FullName attribute
    Set xAttribute = xDoc.createAttribute("FullName")
    If Trim$(FullName) <> "" Then
        xAttribute.Text = Trim$(FullName)
    Else
        'User didn't supply Object Full Name, so use the Chassis.FullName
        xAttribute.Text = cTruck.Chassis.FullName
    End If
    xParent.Attributes.setNamedItem xAttribute
    
    ' create DisplayName attribute if user has defined one
    If Trim$(DisplayName) <> "" Then
        Set xAttribute = xDoc.createAttribute("DisplayName")
        xAttribute.Text = Trim$(DisplayName)
        xParent.Attributes.setNamedItem xAttribute
    End If
    
    ' add the element
    xDoc.appendChild xParent
    
    'Add sub-nodes
    sNodeName = Split("Chassis|Components", "|")
    For i% = 0 To UBound(sNodeName)
        Select Case sNodeName(i%)
        Case "Chassis"
            sErrMsg = FillChassisNode(xDoc, cTruck)
        Case "Components"
            If colComponents.Count > 0 Then
                sErrMsg = FillComponentsNode(xDoc, cTruck, colComponents)
            End If
        End Select
    Next i%
    
    xDoc.save sFile
    SaveChassisObject = "" 'No error if we reached this point

CleanUp:
    Set xDoc = Nothing
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modALObjects.SaveChassisObject(cTruck)", cTruck
    Resume CleanUp
End Function


Private Function ValidObjectFileName(sPrefix As String, sFile As String) As String
    Dim sMsg As String
    Dim bBad As Boolean
    Dim sFileBase As String
    On Error GoTo errHandler
    
    bBad = False
    
    sFileBase = PathlessFileName(sFile)
    If Len(sFileBase) <= Len(sPrefix) Then
        bBad = True
    ElseIf UCase$(Left$(sFileBase, Len(sPrefix))) <> UCase$(sPrefix) Then
        bBad = True
    ElseIf UCase$(Right$(sFileBase, 4)) <> ".XML" Then
        bBad = True
    End If
    
    If bBad Then
        sMsg = "Filenames must beging with '" & sPrefix & "' and end with '.XML'" & vbCrLf & _
               "Please entear a valid name"
    Else
        sMsg = ""
    End If
    
    ValidObjectFileName = sMsg
    Exit Function
errHandler:
    ErrorIn "modALObjects.ValidObjectFileName(sPrefix,sFile)", Array(sPrefix, sFile)
End Function


Public Function ComponentCopy(cComponent As clsComponent) As clsComponent
    'This function returns a COPY of the supplied component, not just a reference
    ' to the original component object
    Dim cComp As New clsComponent
    Dim cRelationship As New clsFillRelationship
    Dim vKey As Variant
    Dim i%
    On Error GoTo errHandler
    
    'Copy top-level properties
    cComp.ContentsType = cComponent.ContentsType
    cComp.DisplayName = cComponent.DisplayName
    cComp.EmptyCG = cComponent.EmptyCG
    cComp.EmptyWeight = cComponent.EmptyWeight
    cComp.FullName = cComponent.FullName
    cComp.InstallationNotes = cComponent.InstallationNotes
    cComp.LocationReference = cComponent.LocationReference
    cComp.Offset = cComponent.Offset
    cComp.PlacementAllowable = cComponent.PlacementAllowable
    cComp.Placement = cComponent.Placement
    cComp.CurbSideStd = cComponent.CurbSideStd
    cComp.StreetSideStd = cComponent.StreetSideStd
    
    
    'Copy RelationShip Objects
    If TypeName(cComponent.FillRelationShips) = "Nothing" Then
        Set cComp.FillRelationShips = Nothing
    Else
        Set cComp.FillRelationShips = New clsFillRelationships
        For Each cRelationship In cComponent.FillRelationShips
            cComp.FillRelationShips.Add cRelationship
        Next
    End If
    
    'Copy Capacity Object
    If TypeName(cComponent.Capacity) = "Nothing" Then
        Set cComp.Capacity = Nothing
    Else
        Set cComp.Capacity = New clsCapacity
        cComp.Capacity.CurContentCG = cComponent.Capacity.CurContentCG
        cComp.Capacity.CurStkHt = cComponent.Capacity.CurStkHt
        cComp.Capacity.CurVol = cComponent.Capacity.CurVol
        cComp.Capacity.DefaultVolContents = cComponent.Capacity.DefaultVolContents
        cComp.Capacity.DefaultVolContents = cComponent.Capacity.DefaultVolContents
        cComp.Capacity.DefaultWtContents = cComponent.Capacity.DefaultWtContents
        cComp.Capacity.DensityContents = cComponent.Capacity.DensityContents
        cComp.Capacity.MaxHt = cComponent.Capacity.MaxHt
        cComp.Capacity.UsesSightGauge = cComponent.Capacity.UsesSightGauge
        cComp.Capacity.Volume = cComponent.Capacity.Volume
    End If
    
    If TypeName(cComponent.Capacity.StickLength) = "Nothing" Then
        Set cComp.Capacity.StickLength = Nothing
    Else
        Set cComp.Capacity.StickLength = New Collection
        For i% = 0 To cComponent.Capacity.StickLength.Count - 1
            vKey = cComponent.Capacity.StickLength(CStr(i%))
            cComp.Capacity.StickLength.Add vKey, CStr(i%)
        Next
    End If
    
    If TypeName(cComponent.Capacity.ContentCG) = "Nothing" Then
        Set cComp.Capacity.ContentCG = Nothing
    Else
        Set cComp.Capacity.ContentCG = New Collection
        For i% = 0 To cComponent.Capacity.ContentCG.Count - 1
            vKey = cComponent.Capacity.ContentCG(CStr(i%))
            cComp.Capacity.ContentCG.Add vKey, CStr(i%)
        Next
    End If
    
CleanUp:
    Set ComponentCopy = cComp
    Set cComp = Nothing
    Set cRelationship = Nothing
    Exit Function
errHandler:
    ErrorIn "modALObjects.ComponentCopy(cComponent)", cComponent
    Resume CleanUp
End Function


Public Function SaveComponentObject(cComponent As clsComponent) As String
    'Saves the Component info of given truck to an object file
    Dim xDoc As MSXML2.DOMDocument30
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim xAttribute As MSXML2.IXMLDOMAttribute
    Dim sFile As String
    Dim sMsg As String
    Dim sErrMsg As String
    On Error GoTo errHandler
    
    SaveComponentObject = "Unknown Error in modALObjects:SaveComponentObject" 'default to fail
    
    With frmMain.dlgFile
        sFile = "Component_" & cComponent.DisplayName & ".xml"
        sFile = FixedFileName(sFile) 'remove any invalid characters
        .FileName = sFile
        .Filter = "Component Files|Component*.xml"
        .DialogTitle = "Save New Component"
        .CancelError = True
        .InitDir = cGlobalInfo.ALObjectFolder
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        On Error Resume Next
        .ShowSave 'Show the SaveFile dialog
        'If User selected a file, open it
        If Err <> 0 Then
            'User canceled
            SaveComponentObject = "File save was aborted."
            GoTo CleanUp
        End If
        On Error GoTo errHandler
        sFile = Trim$(.FileName)
        
        If ValidObjectFileName("Component_", sFile) <> "" Then
            'bad file name
            MsgBox ValidObjectFileName("Component_", sFile), vbCritical, "Error"
            'recursive call to give user 2nd chance
            SaveComponentObject = SaveComponentObject(cComponent)
        ElseIf FileExists(sFile) Then
            'See if user wants to over-write existing file
            sMsg = "The chosen filename exists.  " & vbCrLf & _
                   "Choose 'Yes' to overwrite existing file," & vbCrLf & _
                   "or 'No' to choose another name"
            If MsgBox(sMsg, vbYesNo, "Overwrite existing file?") = vbNo Then
                'recursive call to give user 2nd chance
                SaveComponentObject = SaveComponentObject(cComponent)
                Exit Function
            End If
        End If
    End With
    
    'Now save to indicated file ============================================
    'Create new DOM Document
    Set xDoc = New MSXML2.DOMDocument30
    With xDoc
        .async = False
        .validateOnParse = False
        .resolveExternals = False
        .preserveWhiteSpace = True
    End With
    'create processing instructions
    Set xNode = xDoc.createProcessingInstruction("xml", "version='1.0'")
    xDoc.appendChild xNode
    
    'Create Base Node & attributes
    Set xParent = xDoc.createElement("ALObject")
    ' create ObjectType attribute
    Set xAttribute = xDoc.createAttribute("ObjectType")
    xAttribute.Text = "Component"
    xParent.Attributes.setNamedItem xAttribute
    ' create FullName attribute
    Set xAttribute = xDoc.createAttribute("FullName")
    xAttribute.Text = cComponent.FullName
    xParent.Attributes.setNamedItem xAttribute
    ' add the element
    xDoc.appendChild xParent
    
    Set xNode = xDoc.createElement("Component")
    xParent.appendChild xNode
    sErrMsg = FillComponentNode(xDoc, xNode, cComponent)
    
    xDoc.save sFile
    SaveComponentObject = "" 'No error if we reached this point

CleanUp:
    Set xDoc = Nothing
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modALObjects.SaveComponentObject(cComponent)", cComponent
    Resume CleanUp
End Function

Public Function SaveTruckAs(cTruck As clsTruck, strCurFile As String) As String
    'Save the modified truck to a new file
    Dim sMsg As String
    Dim strNewFile As String
    Dim sFile As String
    On Error GoTo errHandler
    
    'Prevent saving of files created on newer software
    If cTruck.CreateVersion.Major = App.Major And _
           cTruck.CreateVersion.Minor > App.Minor Then
        sMsg = "This file was created on a newer version of this software. " & vbCrLf & _
                "Please contact Tread about upgrading to the latest version."
        MsgBox sMsg, vbExclamation, "Cannot save file"
        Exit Function
    End If
    
    
    If IsNumber(cTruck.SN) Then
        'User (properly) eneterd a number or .SN property
        sFile = "Truck_SN" & Trim$(cTruck.SN) & ".xml"
    Else
        If Len(cTruck.SN) > 2 Then
            If Mid$(cTruck.SN, 1, 1) = "q" Then
                sFile = "Truck_Q" & Mid$(cTruck.SN, 2) & ".xml"
            ElseIf UCase$(Mid$(cTruck.SN, 1, 2)) = "SN" Then
                sFile = "Truck_SN" & Mid$(cTruck.SN, 3) & ".xml"
            Else
                sFile = "Truck_" & Trim$(cTruck.SN) & ".xml"
            End If
        Else
            sFile = "Truck_" & Trim$(cTruck.SN) & ".xml"
        End If
    End If
    
    With frmMain.dlgFile
        strNewFile = ""
        .FileName = sFile 'strCurFile
        .Filter = "Truck Files|Truck*.xml"
        .DialogTitle = "Save New Component"
        .CancelError = True
        .InitDir = cGlobalInfo.TruckFolder
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        On Error Resume Next
        .ShowSave 'Show the SaveFile dialog
        'If User selected a file, open it
        strNewFile = Trim$(.FileName)
        If Err <> 0 Or strNewFile = "" Then
            'User canceled
            SaveTruckAs = "" 'No error for user abort
            Exit Function
        End If
        On Error GoTo errHandler
        strNewFile = Trim$(.FileName)
        
        If ValidObjectFileName("Truck_", strNewFile) <> "" Then
            'bad file name
            MsgBox ValidObjectFileName("Truck_", sFile), vbCritical, "Error"
            'recursive call to give user 2nd chance
            SaveTruckAs = SaveTruckAs(cTruck, sFile)
            Exit Function
        ElseIf FileExists(strNewFile) Then
            'See if user wants to over-write existing file
            sMsg = "The chosen filename exists.  " & vbCrLf & _
                   "Choose 'Yes' to overwrite existing file," & vbCrLf & _
                   "or 'No' to choose another name"
            If MsgBox(sMsg, vbYesNo, "Overwrite existing file?") = vbNo Then
                'recursive call to give user 2nd chance
                    SaveTruckAs = SaveTruckAs(cTruck, sFile)
                    Exit Function
                Exit Function
            End If
        End If
    End With
    
    sMsg = SaveTruck(strNewFile, cTruck)
    If sMsg <> "" Then
        SaveTruckAs = "Failed to Save Truck File"
        Exit Function
    End If
    
    strCurFile = strNewFile
    bHasBeenEdited = False
    bTruckFileDirty = False
    Exit Function
errHandler:
    ErrorIn "modALObjects.SaveTruckAs(cTruck,strCurFile)", Array(cTruck, strCurFile)
End Function

Private Function RemoveChassisComponents(cTruck As clsTruck, sGUID As String)
    Dim i%
    
    If Len(sGUID) < 5 Then Exit Function
    
    'first DELETE components related to existing chassis
    For i% = 1 To cTruck.Components.Count
        If cTruck.Components(i%).ChassisGUID = sGUID Then
            cTruck.Components.Remove i%
            RemoveChassisComponents cTruck, sGUID
            Exit For
        End If
    Next
End Function


Public Function SetDefaultPlacement(plAllowable As PlacementAllowableLocation, plNow As PlacementLocation) As PlacementLocation
    'This function sets/changes .Placement so it doesn't conflict .PlacementAllowable
    Select Case plAllowable
    Case paEitherSide
        If plNow = plCurbSideStd Or plNow = plStreetSideStd Then
            'Leave alone
            SetDefaultPlacement = plNow
        Else
            SetDefaultPlacement = plStreetSideStd
        End If
    Case paStreetSideStd
            SetDefaultPlacement = plStreetSideStd
    Case paCurbSideStd
            SetDefaultPlacement = plCurbSideStd
    Case paCenter
            SetDefaultPlacement = plCenter
    End Select
    
End Function
