Attribute VB_Name = "modXML"
Option Explicit

Public cGlobalInfo As clsGlobalInfo


Public Function LoadTruck(strFile As String, cTruck As clsTruck) As String
    Dim xDoc As MSXML2.DOMDocument30
    On Error GoTo errHandler
    
    LoadTruck = "Unknown Error in modXML:LoadTruck" 'default to fail
    Set xDoc = New MSXML2.DOMDocument30
    xDoc.async = False 'Don't load asynchronysly
    If xDoc.Load(strFile) Then
       ' The document loaded successfully.
       LoadTruck = ""
       LoadTruck = ParseDoc(xDoc, cTruck)
    Else
       ' The document failed to load.
       Dim strErrText As String
       Dim xPE As MSXML2.IXMLDOMParseError
       ' Obtain the ParseError object
       Set xPE = xDoc.parseError
       With xPE
          strErrText = "Your XML Document failed to load" & _
            "due the following error." & vbCrLf & _
            "Error #: " & .ErrorCode & ": " & xPE.reason & _
            "Line #: " & .Line & vbCrLf & _
            "Line Position: " & .linepos & vbCrLf & _
            "Position In File: " & .filepos & vbCrLf & _
            "Source Text: " & .srcText & vbCrLf & _
            "Document URL: " & .url
        End With
        Set xPE = Nothing
        LoadTruck = strErrText
    End If
CleanUp:
    Set xDoc = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.LoadTruck(strFile,cTruck)", Array(strFile, cTruck)
    Resume CleanUp
End Function

Private Function ParseDoc(ByRef xDoc As MSXML2.DOMDocument30, ByRef cTruck As clsTruck) As String
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xChildNodes As MSXML2.IXMLDOMNodeList
    Dim sErrMsg As String
    On Error GoTo errHandler

    ParseDoc = "Unknown Error in modXML:ParseDoc"
    Set xNode = xDoc.selectSingleNode("Truck")
    If TypeName(xNode) = "Nothing" Then
        ParseDoc = "XML File Error.  'Truck' node does not exist"
        Exit Function
    End If
    Set xNodes = xNode.childNodes
    Set cTruck.CreateVersion = New clsVersion 'default (empty)
    Set cTruck.Components = New clsComponents 'default (empty)
    
    For Each xNode In xNodes
        Select Case UCase(xNode.nodeName)
            Case "CREATEVERSION"
                cTruck.CreateVersion.TextVal = xNode.firstChild.nodeValue
            Case "SN"
                cTruck.SN = xNode.firstChild.nodeValue
            Case "OWNER"
                cTruck.Owner = xNode.firstChild.nodeValue
            Case "DESCRIPTION"
                cTruck.Description = xNode.firstChild.nodeValue
            Case "ISSTANDARDMOUNT"
                cTruck.IsStandardMount = CBool(xNode.firstChild.nodeValue)
            Case "BODYLOCATION"
                cTruck.BodyLocation = CDbl(xNode.firstChild.nodeValue)
            Case "BRIDGELAW"
                cTruck.BridgeLaw = xNode.firstChild.nodeValue
            Case "WTADJUSTFRONT"
                cTruck.WtAdjustFront = xNode.firstChild.nodeValue
            Case "WTADJUSTREAR"
                cTruck.WtAdjustRear = xNode.firstChild.nodeValue
            Case "SHIPDATE"
                cTruck.ShipDate = CDate(xNode.firstChild.nodeValue)
            Case "BODY"
                'read child nodes
                Set xChildNodes = xNode.childNodes
                sErrMsg = ParseBody(xChildNodes, cTruck)
                If sErrMsg <> "" Then
                    ParseDoc = sErrMsg
                    Exit Function
                End If
            Case "CHASSIS"
                'read child nodes
                Set xChildNodes = xNode.childNodes
                sErrMsg = ParseChassis(xChildNodes, cTruck)
                If sErrMsg <> "" Then
                    ParseDoc = sErrMsg
                    Exit Function
                End If
            Case "COMPONENTS"
                'read child nodes
                Set xChildNodes = xNode.childNodes
                'read components w/ locations as dist from front axle
                sErrMsg = ParseComponents(xChildNodes, cTruck)
                If sErrMsg <> "" Then
                    ParseDoc = sErrMsg
                    GoTo CleanUp
                End If
        End Select
    Next
    ParseDoc = "" 'Got this far, so return w/o error
CleanUp:
    Set xNode = Nothing
    Set xNodes = Nothing
    Set xChildNodes = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.ParseDoc(xDoc,cTruck)", Array(xDoc, cTruck)
End Function

Public Function ParseBody(ByRef xNodes As MSXML2.IXMLDOMNodeList, ByRef cTruck As clsTruck) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xChildNodes As MSXML2.IXMLDOMNodeList
    Dim sErrMsg As String
    Dim sVal As String
    On Error GoTo errHandler

    ParseBody = "Unknown error in modXML:ParseBody"
    Set cTruck.Body = New clsBody
    Set cTruck.Body.Tanks = New clsTanks 'default (empty)
    For Each xNode In xNodes
        Select Case UCase(xNode.nodeName)
            Case "FULLNAME"
                cTruck.Body.FullName = xNode.firstChild.nodeValue
            Case "DISPLAYNAME"
                cTruck.Body.DisplayName = xNode.firstChild.nodeValue
            Case "EMPTYWEIGHT"
                cTruck.Body.EmptyWeight = CDbl(xNode.firstChild.nodeValue)
            Case "EMPTYCG"
                cTruck.Body.EmptyCG = CDbl(xNode.firstChild.nodeValue)
            Case "BODYLENGTH"
                cTruck.Body.BodyLength = CDbl(xNode.firstChild.nodeValue)
            Case "PLACEMENTALLOWABLE"
                sVal = xNode.firstChild.nodeValue
                Select Case UCase$(Trim$(sVal))
                Case "EITHER SIDE"
                    cTruck.Body.PlacementAllowable = paEitherSide
                Case "STREETSIDESTD"
                    cTruck.Body.PlacementAllowable = paStreetSideStd
                Case "CURBSIDESTD"
                    cTruck.Body.PlacementAllowable = paCurbSideStd
                Case "CENTER"
                    cTruck.Body.PlacementAllowable = paCenter
                End Select
            Case "PLACEMENT"
                sVal = xNode.firstChild.nodeValue
                Select Case UCase$(Trim$(sVal))
                Case "NOT PLACED", ""
                    cTruck.Body.Placement = plNotPlaced
                Case "STREETSIDESTD"
                    cTruck.Body.Placement = plStreetSideStd
                Case "CURBSIDESTD"
                    cTruck.Body.Placement = plCurbSideStd
                Case "CENTER"
                    cTruck.Body.Placement = plCenter
                End Select
            Case "STREETSIDESTD"
                cTruck.Body.StreetSideStd = ""
                If xNode.hasChildNodes Then cTruck.Body.StreetSideStd = xNode.firstChild.nodeValue
            Case "CURBSIDESTD"
                cTruck.Body.CurbSideStd = ""
                If xNode.hasChildNodes Then cTruck.Body.CurbSideStd = xNode.firstChild.nodeValue
            Case "TANKS"
                'read child nodes
                Set xChildNodes = xNode.childNodes
                sErrMsg = ParseTanks(xChildNodes, cTruck)
                If sErrMsg <> "" Then
                    ParseBody = sErrMsg
                    GoTo CleanUp
                End If
        End Select
    Next
    ParseBody = "" 'Got this far, so return w/o error
CleanUp:
    Set xNode = Nothing
    Set xChildNodes = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.ParseBody(xNodes,cTruck)", Array(xNodes, cTruck)
    Resume CleanUp
End Function

Private Function ParseTanks(xTankNodes As MSXML2.IXMLDOMNodeList, ByRef cTruck As clsTruck) As String
    Dim xTank As MSXML2.IXMLDOMNode
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xChildNodes As MSXML2.IXMLDOMNodeList
    Dim xChildNode As MSXML2.IXMLDOMNode
    Dim cTank As clsTank
    Dim sName As String
    Dim sVal As String
    On Error GoTo errHandler

    ParseTanks = "Unknown error in modXML:ParseTanks"
    Set cTruck.Body.Tanks = New clsTanks
    For Each xTank In xTankNodes
        Set cTank = New clsTank
        Set cTank.StickLength = New Collection 'default (empty)
        Set cTank.ContentCG = New Collection 'default (empty)
        Set xNodes = xTank.childNodes
        For Each xNode In xNodes
            Select Case UCase(xNode.nodeName)
                Case "DISPLAYNAME"
                    cTank.DisplayName = xNode.firstChild.nodeValue
                Case "TANKTYPE"
                    sVal = xNode.firstChild.nodeValue
                    Select Case UCase(sVal)
                    Case "AN"
                        cTank.TankType = ttAN
                    Case "EMULSION"
                        cTank.TankType = ttEmulsion
                    Case "DUAL"
                        cTank.TankType = ttDual
                    End Select
                Case "CURTANKUSE"
                    sVal = xNode.firstChild.nodeValue
                    Select Case UCase(sVal)
                    Case "EMULSION"
                        cTank.CurTankUse = ttEmulsion
                    Case Else
                        cTank.CurTankUse = ttAN
                    End Select
                Case "VOLUME"
                    cTank.Volume = CDbl(xNode.firstChild.nodeValue)
                Case "MAXMASS"
                    cTank.MaxMass = xNode.firstChild.nodeValue
                Case "MAXMASSDESC"
                    cTank.MaxMassDesc = xNode.firstChild.nodeValue
                Case "STICKLENGTH"
                    'read child nodes into collection
                    Set xChildNodes = xNode.childNodes
                    Set cTank.StickLength = New Collection
                    For Each xChildNode In xChildNodes
                        sName = xChildNode.nodeName
                        sName = Right$(sName, 1) 'turns "K0" into "0"
                        sVal = xChildNode.firstChild.nodeValue
                        cTank.StickLength.Add CDbl(sVal), sName
                    Next
                Case "CONTENTCG"
                    'read child nodes into collection
                    Set xChildNodes = xNode.childNodes
                    Set cTank.ContentCG = New Collection
                    For Each xChildNode In xChildNodes
                        sName = xChildNode.nodeName
                        sName = Right$(sName, 1) 'turns "K0" into "0"
                        sVal = xChildNode.firstChild.nodeValue
                        cTank.ContentCG.Add CDbl(sVal), sName
                    Next
                Case "CONTENTVCG"
                    'read child nodes into collection
                    Set xChildNodes = xNode.childNodes
                    Set cTank.ContentVCG = New Collection
                    For Each xChildNode In xChildNodes
                        sName = xChildNode.nodeName
                        sName = Right$(sName, 1) 'turns "K0" into "0"
                        sVal = xChildNode.firstChild.nodeValue
                        cTank.ContentVCG.Add CDbl(sVal), sName
                    Next
                End Select
        Next
        cTruck.Body.Tanks.Add cTank
    Next
    ParseTanks = "" 'Got this far, so return w/o error
    
CleanUp:
    Set cTank = Nothing
    Set xTank = Nothing
    Set xNodes = Nothing
    Set xNode = Nothing
    Set xChildNodes = Nothing
    Set xChildNode = Nothing
    Exit Function
    
errHandler:
    ErrorIn "modXML.ParseTanks(xNodes,cTruck)", Array(xNodes, cTruck)
    Resume CleanUp
End Function

Public Function ParseChassis(xNodes As MSXML2.IXMLDOMNodeList, ByRef cTruck As clsTruck) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xChildNodes As MSXML2.IXMLDOMNodeList
    Dim sErrMsg As String
    Dim sVal As String
    On Error GoTo errHandler

    ParseChassis = "Unknown error in modXML:ParseChassis"
    Set cTruck.Chassis = New clsChassis
    Set cTruck.Chassis.Tags = New clsTags 'default (empty)
    For Each xNode In xNodes
        Select Case UCase(xNode.nodeName)
            Case "FULLNAME"
                cTruck.Chassis.FullName = xNode.firstChild.nodeValue
            Case "DISPLAYNAME"
                cTruck.Chassis.DisplayName = xNode.firstChild.nodeValue
            Case "WB"
                cTruck.Chassis.WB = CDbl(xNode.firstChild.nodeValue)
            Case "TWINSTEERSEPARATION"
                cTruck.Chassis.TwinSteerSeparation = CDbl(xNode.firstChild.nodeValue)
            Case "BACKOFCAB"
                cTruck.Chassis.BackOfCab = CDbl(xNode.firstChild.nodeValue)
            Case "WTFRONT"
                cTruck.Chassis.WtFront = CDbl(xNode.firstChild.nodeValue)
            Case "WTREAR"
                cTruck.Chassis.WtRear = CDbl(xNode.firstChild.nodeValue)
            Case "WTLIMITFRONT"
                cTruck.Chassis.WtLimitFront = CDbl(xNode.firstChild.nodeValue)
            Case "WTLIMITREAR"
                cTruck.Chassis.WtLimitRear = CDbl(xNode.firstChild.nodeValue)
            Case "WTLIMITTOTAL"
                cTruck.Chassis.WtLimitTotal = CDbl(xNode.firstChild.nodeValue)
            Case "TANDEMSPACING"
                cTruck.Chassis.TandemSpacing = CDbl(xNode.firstChild.nodeValue)
            Case "PLACEMENTALLOWABLE"
                sVal = xNode.firstChild.nodeValue
                Select Case UCase$(Trim$(sVal))
                Case "EITHER SIDE"
                    cTruck.Chassis.PlacementAllowable = paEitherSide
                Case "STREETSIDESTD"
                    cTruck.Chassis.PlacementAllowable = paStreetSideStd
                Case "CURBSIDESTD"
                    cTruck.Chassis.PlacementAllowable = paCurbSideStd
                Case "CENTER"
                    cTruck.Chassis.PlacementAllowable = paCenter
                End Select
            Case "PLACEMENT"
                sVal = xNode.firstChild.nodeValue
                Select Case UCase$(Trim$(sVal))
                Case "NOT PLACED", ""
                    cTruck.Chassis.Placement = plNotPlaced
                Case "STREETSIDESTD"
                    cTruck.Chassis.Placement = plStreetSideStd
                Case "CURBSIDESTD"
                    cTruck.Chassis.Placement = plCurbSideStd
                Case "CENTER"
                    cTruck.Chassis.Placement = plCenter
                End Select
            Case "STREETSIDESTD"
                cTruck.Chassis.StreetSideStd = ""
                If xNode.hasChildNodes Then cTruck.Chassis.StreetSideStd = xNode.firstChild.nodeValue
            Case "CURBSIDESTD"
                cTruck.Chassis.CurbSideStd = ""
                If xNode.hasChildNodes Then cTruck.Chassis.CurbSideStd = xNode.firstChild.nodeValue
            Case "WHEELDIA"
                cTruck.Chassis.WheelDia = 45 * 0.0254 'Default Value
                If xNode.hasChildNodes Then cTruck.Chassis.WheelDia = CDbl(xNode.firstChild.nodeValue)
            Case "WHEELY"
                cTruck.Chassis.WheelY = -22 * 0.0254 'Default Value
                If xNode.hasChildNodes Then cTruck.Chassis.WheelY = CDbl(xNode.firstChild.nodeValue)
            Case "FRAMERAILHT"
                cTruck.Chassis.FrameRailHt = 12 * 0.0254 'Default Value
                If xNode.hasChildNodes Then cTruck.Chassis.FrameRailHt = CDbl(xNode.firstChild.nodeValue)
            Case "TAGS"
                'read child nodes
                Set xChildNodes = xNode.childNodes
                sErrMsg = ParseTags(xChildNodes, cTruck)
                If sErrMsg <> "" Then
                    ParseChassis = sErrMsg
                    GoTo CleanUp
                End If
        End Select
    Next
    If cTruck.Chassis.WtLimitTotal = 0 Then
        'In case there was no specified gross wt limit, define by sum of axle limits
        cTruck.Chassis.WtLimitTotal = cTruck.Chassis.WtLimitFront + cTruck.Chassis.WtLimitRear
    End If
    
    cTruck.Chassis.GUID = CreateGUID 'Assign a GUID so imported components can be linked to the chassis
    ParseChassis = "" 'Got this far, so return w/o error
CleanUp:
    Set xNode = Nothing
    Set xChildNodes = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.ParseChassis(xNodes,cTruck)", Array(xNodes, cTruck)
    Resume CleanUp
End Function

Private Function ParseTags(xTags As MSXML2.IXMLDOMNodeList, ByRef cTruck As clsTruck) As String
    Dim xTag As MSXML2.IXMLDOMNode
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim cTag As clsTag
    On Error GoTo errHandler

    ParseTags = "Unknown error in modXML:ParseTags"
    Set cTruck.Chassis.Tags = New clsTags
    For Each xTag In xTags
        Set xNodes = xTag.childNodes
        Set cTag = New clsTag
        For Each xNode In xNodes
            Select Case UCase(xNode.nodeName)
                Case "LOCATION"
                    cTag.Location = CDbl(xNode.firstChild.nodeValue)
                Case "WEIGHT"
                    cTag.Weight = CDbl(xNode.firstChild.nodeValue)
                Case "WTLIMIT"
                    cTag.WtLimit = CDbl(xNode.firstChild.nodeValue)
                Case "FORCETOPRESSURE"
                    cTag.ForceToPressure = CDbl(xNode.firstChild.nodeValue)
                Case "DOWNWARDFORCE"
                    cTag.DownwardForce = CDbl(xNode.firstChild.nodeValue)
            End Select
        Next
        cTruck.Chassis.Tags.Add cTag
    Next
    ParseTags = "" 'Got this far, so return w/o error
CleanUp:
    Set xTag = Nothing
    Set xNodes = Nothing
    Set xNode = Nothing
    Set cTag = Nothing
    Exit Function
    
errHandler:
    ErrorIn "modXML.ParseTags(xTags,cTruck)", Array(xTags, cTruck)
    Resume CleanUp
End Function

Public Function ParseComponents(xComponents As MSXML2.IXMLDOMNodeList, ByRef cTruck As clsTruck, Optional bLinkToChassis As Boolean = False) As String
    On Error GoTo errHandler
    'Reads the Components object.
    'The "LocationReference" variable is used to identify the Origin Reference of
    ' the components that will be read. Locations can be referenced to either the
    'front of the body (default), or to the front axle.
    Dim xComponent As MSXML2.IXMLDOMNode
    Dim errMsg As String
    Dim sRtn As String
    
    errMsg = ""
    If TypeName(cTruck.Components) = "Nothing" Then
        'Only need to do this if Components object doesn't already exist
        Set cTruck.Components = New clsComponents
    End If
    
    For Each xComponent In xComponents
        sRtn = ParseComponent(xComponent, cTruck, , bLinkToChassis)
        If sRtn <> "" Then
            errMsg = errMsg & "- " & sRtn & vbCrLf
        End If
    Next
    ParseComponents = errMsg 'return any error messages
    
CleanUp:
    Set xComponent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.ParseComponents(xComponents,cTruck)", Array(xComponents, cTruck)
    ParseComponents = errMsg & "- " & Err.Description
    Resume CleanUp
End Function


Public Function ParseComponent(xComponent As MSXML2.IXMLDOMNode, ByRef cTruck As clsTruck, Optional LocationReference As OriginReference = orBodyOrigin, Optional bLinkToChassis As Boolean = False) As String
    Dim cComponent As clsComponent
    Dim xChildNodes As MSXML2.IXMLDOMNodeList
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim sVal As String
    Dim sErrMsg As String
    On Error GoTo errHandler
    
    Set cComponent = New clsComponent
    Set cComponent.FillRelationShips = New clsFillRelationships 'default (empty)
    Set cComponent.Capacity = New clsCapacity 'default (empty)
    Set xNodes = xComponent.childNodes
    cComponent.LocationReference = LocationReference 'set default value
    For Each xNode In xNodes
        Select Case UCase(xNode.nodeName)
            Case "LOCATIONREFERENCE"
                Select Case UCase(xNode.firstChild.nodeValue)
                Case "BODY"
                    cComponent.LocationReference = orBodyOrigin
                Case "CHASSIS", "FRONTAXLE"
                    cComponent.LocationReference = orFrontAxle
                End Select
            Case "FULLNAME"
                cComponent.FullName = xNode.firstChild.nodeValue
            Case "DISPLAYNAME"
                cComponent.DisplayName = xNode.firstChild.nodeValue
            Case "OFFSET"
                cComponent.Offset = CDbl(xNode.firstChild.nodeValue)
            Case "EMPTYWEIGHT"
                cComponent.EmptyWeight = CDbl(xNode.firstChild.nodeValue)
            Case "EMPTYCG"
                cComponent.EmptyCG = CDbl(xNode.firstChild.nodeValue)
            Case "CONTENTSTYPE"
                sVal = xNode.firstChild.nodeValue
                Select Case UCase(sVal)
                Case "NONE"
                    cComponent.ContentsType = ctNone
                Case "FUEL"
                    cComponent.ContentsType = ctFuel
                Case "WATER"
                    cComponent.ContentsType = ctWater
                Case "GASA"
                    cComponent.ContentsType = ctGasA
                Case "GASB"
                    cComponent.ContentsType = ctGasB
                Case "ADDITIVE"
                    cComponent.ContentsType = ctAdditive
                Case "OTHER"
                    cComponent.ContentsType = ctOther
                End Select
            Case "PLACEMENTALLOWABLE"
                sVal = xNode.firstChild.nodeValue
                Select Case UCase$(Trim$(sVal))
                Case "EITHER SIDE"
                    cComponent.PlacementAllowable = paEitherSide
                Case "STREETSIDESTD"
                    cComponent.PlacementAllowable = paStreetSideStd
                Case "CURBSIDESTD"
                    cComponent.PlacementAllowable = paCurbSideStd
                Case "CENTER"
                    cComponent.PlacementAllowable = paCenter
                End Select
            Case "PLACEMENT"
                sVal = xNode.firstChild.nodeValue
                Select Case UCase$(Trim$(sVal))
                Case "NOT PLACED", ""
                    cComponent.Placement = plNotPlaced
                Case "STREETSIDESTD"
                    cComponent.Placement = plStreetSideStd
                Case "CURBSIDESTD"
                    cComponent.Placement = plCurbSideStd
                Case "CENTER"
                    cComponent.Placement = plCenter
                End Select
            Case "STREETSIDESTD"
                cComponent.StreetSideStd = ""
                If xNode.hasChildNodes Then cComponent.StreetSideStd = xNode.firstChild.nodeValue
            Case "CURBSIDESTD"
                cComponent.CurbSideStd = ""
                If xNode.hasChildNodes Then cComponent.CurbSideStd = xNode.firstChild.nodeValue
            Case "FILLRELATIONSHIPS"
                'read child nodes
                Set xChildNodes = xNode.childNodes
                sErrMsg = ParseRelationships(xChildNodes, cComponent)
                If sErrMsg <> "" Then
                    ParseComponent = sErrMsg
                    GoTo CleanUp
                End If
            
            Case "CAPACITY"
                'read child nodes
                Set xChildNodes = xNode.childNodes
                sErrMsg = ParseCapacity(xChildNodes, cComponent)
                If sErrMsg <> "" Then
                    ParseComponent = sErrMsg
                    GoTo CleanUp
                End If
            
            Case "INSTALLATIONNOTES"
                If TypeName(xNode.firstChild) <> "Nothing" Then
                    cComponent.InstallationNotes = xNode.firstChild.nodeValue
                End If
        End Select
    Next
    If bLinkToChassis Then cComponent.ChassisGUID = cTruck.Chassis.GUID
    
    'Make sure component have a valid value for .Placement
    cComponent.Placement = SetDefaultPlacement(cComponent.PlacementAllowable, cComponent.Placement)
    
    cTruck.Components.Add cComponent

CleanUp:
    Set cComponent = Nothing
    Set xChildNodes = Nothing
    Set xNodes = Nothing
    Set xNode = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.ParseComponent(xComponent,cTruck,LocationReference)", Array(xComponent, cTruck, _
         LocationReference)
    Resume CleanUp
End Function

Private Function ParseRelationships(xRelationships As MSXML2.IXMLDOMNodeList, ByRef cComponent As clsComponent) As String
    Dim xRelationship As MSXML2.IXMLDOMNode
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim cRelationship As clsFillRelationship
    Dim sVal As String
    On Error GoTo errHandler

    ParseRelationships = "Unknown error in modXML:ParseRelationships"
    
    Set cComponent.FillRelationShips = New clsFillRelationships
    For Each xRelationship In xRelationships
        Set cRelationship = New clsFillRelationship
        Set xNodes = xRelationship.childNodes
        For Each xNode In xNodes
            Select Case UCase(xNode.nodeName)
                Case "PARENTPRODUCT"
                    sVal = xNode.firstChild.nodeValue
                    Select Case UCase(sVal)
                    Case "AN"
                        cRelationship.ParentProduct = ptAN
                    Case "EMULSION", "EMUL"
                        cRelationship.ParentProduct = ptEmulsion
                    End Select
                Case "MULTIPLIER"
                    cRelationship.Multiplier = CDbl(xNode.firstChild.nodeValue)
                Case "OFFSET"
                    cRelationship.Offset = CDbl(xNode.firstChild.nodeValue)
            End Select
        Next
        cComponent.FillRelationShips.Add cRelationship
    Next
    ParseRelationships = "" 'Got this far, so return w/o error

CleanUp:
    Set xRelationship = Nothing
    Set xNodes = Nothing
    Set xNode = Nothing
    Set cRelationship = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.ParseRelationships(xRelationships,cComponent)", Array(xRelationships, cComponent)
    Resume CleanUp
End Function

                    
Private Function ParseCapacity(xNodes As MSXML2.IXMLDOMNodeList, ByRef cComponent As clsComponent) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xChildNodes As MSXML2.IXMLDOMNodeList
    Dim xChildNode As MSXML2.IXMLDOMNode
    Dim sName As String
    Dim sVal As String
    On Error GoTo errHandler

    ParseCapacity = "Unknown error in modXML:ParseCapacity"
    
    Set cComponent.Capacity = New clsCapacity
    Set cComponent.Capacity.StickLength = New Collection 'default (empty)
    Set cComponent.Capacity.ContentCG = New Collection 'default (empty)
    For Each xNode In xNodes
        Select Case UCase(xNode.nodeName)
            Case "DEFAULTWTCONTENTS"
                If TypeName(xNode.firstChild) = "Nothing" Then
                    cComponent.Capacity.DefaultWtContents = ""
                Else
                    cComponent.Capacity.DefaultWtContents = Trim(xNode.firstChild.nodeValue)
                End If
            Case "DEFAULTVOLCONTENTS"
                If TypeName(xNode.firstChild) = "Nothing" Then
                    cComponent.Capacity.DefaultVolContents = ""
                Else
                    cComponent.Capacity.DefaultVolContents = Trim(xNode.firstChild.nodeValue)
                End If
            Case "DENSITYCONTENTS"
                cComponent.Capacity.DensityContents = CDbl(xNode.firstChild.nodeValue)
            Case "USESSIGHTGAUGE"
                cComponent.Capacity.UsesSightGauge = CBool(xNode.firstChild.nodeValue)
            Case "VOLUME"
                cComponent.Capacity.Volume = CDbl(xNode.firstChild.nodeValue)
            Case "STICKLENGTH"
                'read child nodes into collection
                Set cComponent.Capacity.StickLength = New Collection
                Set xChildNodes = xNode.childNodes
                For Each xChildNode In xChildNodes
                    sName = xChildNode.nodeName
                    sName = Right$(sName, 1) 'turns "K0" into "0"
                    sVal = xChildNode.firstChild.nodeValue
                    cComponent.Capacity.StickLength.Add CDbl(sVal), sName
                Next
            Case "CONTENTCG"
                'read child nodes into collection
                Set cComponent.Capacity.ContentCG = New Collection
                Set xChildNodes = xNode.childNodes
                For Each xChildNode In xChildNodes
                    sName = xChildNode.nodeName
                    sName = Right$(sName, 1) 'turns "K0" into "0"
                    sVal = xChildNode.firstChild.nodeValue
                    cComponent.Capacity.ContentCG.Add CDbl(sVal), sName
                Next
        End Select
    Next
    ParseCapacity = "" 'Got this far, so return w/o error
    
CleanUp:
    Set xNode = Nothing
    Set xChildNodes = Nothing
    Set xChildNode = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.ParseCapacity(xNodes,cComponent)", Array(xNodes, cComponent)
    Resume CleanUp
End Function


Public Function ReadGlobalInfo(strFile As String) As String
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xDoc As MSXML2.DOMDocument30
    On Error GoTo errHandler
    
    Set cGlobalInfo = New clsGlobalInfo
    
    ReadGlobalInfo = "Unknown Error in modXML:ReadGlobalInfo" 'default to fail
    Set xDoc = New MSXML2.DOMDocument30
    xDoc.async = False 'Don't load asynchronysly
    If Not xDoc.Load(strFile) Then
       ' The document failed to load.
       Dim strErrText As String
       Dim xPE As MSXML2.IXMLDOMParseError
       ' Obtain the ParseError object
       Set xPE = xDoc.parseError
       With xPE
          strErrText = "Your XML Document failed to load" & _
            "due the following error." & vbCrLf & _
            "Error #: " & .ErrorCode & ": " & xPE.reason & _
            "Line #: " & .Line & vbCrLf & _
            "Line Position: " & .linepos & vbCrLf & _
            "Position In File: " & .filepos & vbCrLf & _
            "Source Text: " & .srcText & vbCrLf & _
            "Document URL: " & .url
        End With
        Set xPE = Nothing
        ReadGlobalInfo = strErrText
    Else
       ' The document loaded successfully.
        Set xNode = xDoc.selectSingleNode("GlobalInfo")
        If TypeName(xNode) = "Nothing" Then
            ReadGlobalInfo = "XML File Error.  'GlobalInfo' node does not exist"
            Exit Function
        End If
        Set xNodes = xNode.childNodes
       
        For Each xNode In xNodes
            Select Case UCase(xNode.nodeName)
            Case "DENSITYAN"
                cGlobalInfo.DensityAN = CDbl(xNode.firstChild.nodeValue)
            Case "DENSITYEMUL"
                cGlobalInfo.DensityEmul = CDbl(xNode.firstChild.nodeValue)
            Case "DENSITYFUEL"
                cGlobalInfo.DensityFuel = CDbl(xNode.firstChild.nodeValue)
            Case "DENSITYWATER"
                cGlobalInfo.DensityWater = CDbl(xNode.firstChild.nodeValue)
            Case "DENSITYGASA"
                cGlobalInfo.DensityGasA = CDbl(xNode.firstChild.nodeValue)
            Case "DENSITYGASB"
                cGlobalInfo.DensityGasB = CDbl(xNode.firstChild.nodeValue)
            Case "DENSITYADDITIVE"
                cGlobalInfo.DensityAdditive = CDbl(xNode.firstChild.nodeValue)
            Case "TRUCKFOLDER"
                If TypeName(xNode.firstChild) = "Nothing" Then
                    cGlobalInfo.TruckFolder = ""
                Else
                    cGlobalInfo.TruckFolder = xNode.firstChild.nodeValue
                End If
                cGlobalInfo.TruckFolder = AddBackslash(cGlobalInfo.TruckFolder)
            Case "ALOBJECTFOLDER"
                If TypeName(xNode.firstChild) = "Nothing" Then
                    cGlobalInfo.ALObjectFolder = ""
                Else
                    cGlobalInfo.ALObjectFolder = xNode.firstChild.nodeValue
                End If
                cGlobalInfo.ALObjectFolder = AddBackslash(cGlobalInfo.ALObjectFolder)
            Case "AVOIDLOWFILL"
                cGlobalInfo.AvoidLowFill = CBool(xNode.firstChild.nodeValue)
            Case "VOLUMEUNITS"
                Set cGlobalInfo.VolumeUnits = New clsUnits
                cGlobalInfo.VolumeUnits.Display = xNode.Attributes.getNamedItem("Display").nodeTypedValue
                cGlobalInfo.VolumeUnits.Multiplier = CDbl(xNode.Attributes.getNamedItem("Multiplier").nodeTypedValue)
            Case "DISTANCEUNITS"
                Set cGlobalInfo.DistanceUnits = New clsUnits
                cGlobalInfo.DistanceUnits.Display = xNode.Attributes.getNamedItem("Display").nodeTypedValue
                cGlobalInfo.DistanceUnits.Multiplier = CDbl(xNode.Attributes.getNamedItem("Multiplier").nodeTypedValue)
            Case "MASSUNITS"
                Set cGlobalInfo.MassUnits = New clsUnits
                cGlobalInfo.MassUnits.Display = xNode.Attributes.getNamedItem("Display").nodeTypedValue
                cGlobalInfo.MassUnits.Multiplier = CDbl(xNode.Attributes.getNamedItem("Multiplier").nodeTypedValue)
            End Select
        Next
    End If
    If cGlobalInfo.TruckFolder = "\" Then
        cGlobalInfo.TruckFolder = AddBackslash(App.Path) & "XML Trucks\"
    End If
    ReadGlobalInfo = "" 'no error if we got this far

CleanUp:
    Set xNodes = Nothing
    Set xNode = Nothing
    Set xDoc = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.ReadGlobalInfo(strFile)", strFile
    Resume CleanUp
End Function


Public Function SaveGlobalInfo(strFile As String) As String
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xDoc As MSXML2.DOMDocument30
    On Error GoTo errHandler
    
    SaveGlobalInfo = "Unknown Error in modXML:SaveGlobalInfo" 'default to fail
    Set xDoc = New MSXML2.DOMDocument30
    xDoc.async = False 'Don't load asynchronysly
    If Not xDoc.Load(strFile) Then
       ' The document failed to load.
       Dim strErrText As String
       Dim xPE As MSXML2.IXMLDOMParseError
       ' Obtain the ParseError object
       Set xPE = xDoc.parseError
       With xPE
          strErrText = "Your XML Document failed to load" & _
            "due the following error." & vbCrLf & _
            "Error #: " & .ErrorCode & ": " & xPE.reason & _
            "Line #: " & .Line & vbCrLf & _
            "Line Position: " & .linepos & vbCrLf & _
            "Position In File: " & .filepos & vbCrLf & _
            "Source Text: " & .srcText & vbCrLf & _
            "Document URL: " & .url
        End With
        Set xPE = Nothing
        SaveGlobalInfo = strErrText
        GoTo CleanUp
    Else
       ' The document loaded successfully.
        Set xNode = xDoc.selectSingleNode("GlobalInfo")
        If TypeName(xNode) = "Nothing" Then
            SaveGlobalInfo = "XML File Error.  'GlobalInfo' node does not exist"
            GoTo CleanUp
        End If
        Set xNodes = xNode.childNodes
       
       'Update the values
        For Each xNode In xNodes
            Select Case UCase(xNode.nodeName)
            Case "DENSITYAN"
                xNode.Text = cGlobalInfo.DensityAN
            Case "DENSITYEMUL"
                xNode.Text = cGlobalInfo.DensityEmul
            Case "DENSITYFUEL"
                xNode.Text = cGlobalInfo.DensityFuel
            Case "DENSITYWATER"
                xNode.Text = cGlobalInfo.DensityWater
            Case "DENSITYGASA"
                xNode.Text = cGlobalInfo.DensityGasA
            Case "DENSITYGASB"
                xNode.Text = cGlobalInfo.DensityGasB
            Case "DENSITYADDITIVE"
                xNode.Text = cGlobalInfo.DensityAdditive
            Case "TRUCKFOLDER"
                xNode.Text = cGlobalInfo.TruckFolder
            Case "AVOIDLOWFILL"
                xNode.Text = cGlobalInfo.AvoidLowFill
            Case "VOLUMEUNITS"
                xNode.Attributes.getNamedItem("Display").nodeTypedValue = cGlobalInfo.VolumeUnits.Display
                xNode.Attributes.getNamedItem("Multiplier").nodeTypedValue = cGlobalInfo.VolumeUnits.Multiplier
            Case "DISTANCEUNITS"
                xNode.Attributes.getNamedItem("Display").nodeTypedValue = cGlobalInfo.DistanceUnits.Display
                xNode.Attributes.getNamedItem("Multiplier").nodeTypedValue = cGlobalInfo.DistanceUnits.Multiplier
            Case "MASSUNITS"
                xNode.Attributes.getNamedItem("Display").nodeTypedValue = cGlobalInfo.MassUnits.Display
                xNode.Attributes.getNamedItem("Multiplier").nodeTypedValue = cGlobalInfo.MassUnits.Multiplier
            End Select
        Next
    End If
    xDoc.save strFile
    
    SaveGlobalInfo = "" 'no error if we got this far

CleanUp:
    Set xNodes = Nothing
    Set xNode = Nothing
    Set xDoc = Nothing
    Exit Function

errHandler:
    ErrorIn "modXML.SaveGlobalInfo(strFile)", strFile
    Resume CleanUp
End Function


Public Function SaveTruck(strFile As String, cTruck As clsTruck) As String
    'Saves the supplied
    Dim xDoc As MSXML2.DOMDocument30
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim sErrMsg As String
    Dim i%
    Dim sNodes As String
    Dim sNodeName() As String
    On Error GoTo errHandler
    
    SaveTruck = "Unknown Error in modXML:SaveTruck" 'default to fail
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
    'Create Base Node
    Set xParent = xDoc.createElement("Truck")
    xDoc.appendChild xParent
    
    'Since we're saving the truck file, stamp the file with the current program version
    cTruck.CreateVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    'Add stand-alone Truck Nodes
    sNodes = "CreateVersion|SN|Owner|Description|IsStandardMount|BodyLocation|BridgeLaw|WtAdjustFront|WtAdjustRear|ShipDate"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        Log "Created '" & sNodeName(i%) & "' node." '***********
        xNode.Text = CallByNameEx(cTruck, sNodeName(i%), VbGet)
        Log "Node text = '" & xNode.Text & "'." '************
        If xNode.Text <> "" Then
            xParent.appendChild xNode
        End If
    Next i%
    
    'Add sub-nodes
    sNodeName = Split("Body|Chassis|Components", "|")
    For i% = 0 To UBound(sNodeName)
        Select Case sNodeName(i%)
        Case "Body"
            sErrMsg = FillBodyNode(xDoc, cTruck)
        Case "Chassis"
            sErrMsg = FillChassisNode(xDoc, cTruck)
        Case "Components"
            If cTruck.Components.Count > 0 Then
                sErrMsg = FillComponentsNode(xDoc, cTruck)
            End If
        End Select
    Next i%
    
    xDoc.save strFile
    SaveTruck = "" 'No error if we reached this point

CleanUp:
    Set xDoc = Nothing
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.SaveTruck(strFile,cTruck)", Array(strFile, cTruck)
End Function


Private Function FillBodyNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef cTruck As clsTruck) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim sErrMsg As String
    Dim i%
    Dim sNodeName() As String
    Dim sNodes As String
    On Error GoTo errHandler
    
    FillBodyNode = "Unknown Error in modXML:FillBodyNode"
    'Create Body node
    Set xParent = xDoc.createElement("Body")
    xDoc.selectSingleNode("Truck").appendChild xParent
        
    'Add stand-alone Body Nodes
    sNodes = "FullName|DisplayName|EmptyWeight|EmptyCG|BodyLength|PlacementAllowable|Placement|StreetSideStd|CurbSideStd"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        xNode.Text = CallByNameEx(cTruck.Body, sNodeName(i%), VbGet)
        If sNodeName(i%) = "PlacementAllowable" Then xNode.Text = cTruck.Body.PlacementAllowableString
        If sNodeName(i%) = "Placement" Then xNode.Text = cTruck.Body.PlacementString
        If xNode.Text <> "" Then
            xParent.appendChild xNode
        End If
    Next i%
    
    'Add tank sub-nodes
    Set xNode = xDoc.createElement("Tanks")
    xParent.appendChild xNode
    'Make 'Tanks' node the parent
    Set xParent = xNode
    'Add 'Tank' nodes
    For i% = 1 To cTruck.Body.Tanks.Count
        Set xNode = xDoc.createElement("Tank")
        xParent.appendChild xNode
        sErrMsg = FillTankNode(xDoc, xNode, cTruck.Body.Tanks(i%))
    Next i%
    
    FillBodyNode = "" 'No error if we reached this point
    
CleanUp:
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillBodyNode(xDoc,cTruck)", Array(xDoc, cTruck)
End Function
        
        
Private Function FillTankNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef xTankNode As IXMLDOMNode, cTank As clsTank) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim sNode As String
    Dim sNodes As String
    Dim sNodeName() As String
    Dim i%
    On Error GoTo errHandler
    
    FillTankNode = "Unknown Error in modXML:FillTankNode"
    'Add stand-alone Body Nodes
    sNodes = "DisplayName|TankType|CurTankUse|Volume|MaxMass|MaxMassDesc"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        xNode.Text = CallByNameEx(cTank, sNodeName(i%), VbGet)
        If sNodeName(i%) = "TankType" Then xNode.Text = cTank.TankTypeString
        If sNodeName(i%) = "CurTankUse" Then xNode.Text = cTank.CurTankUseString
        If (sNodeName(i%) = "MaxMass" And val(xNode.Text) = 0) _
         Or (sNodeName(i%) <> "MaxMass" And xNode.Text = "") Then
            'zero values for MaxMass or empty-string value. Do nothing
        Else
            'only save non-zero values for MaxMass and non-empty-string values
            xTankNode.appendChild xNode
        End If
    Next i%
    
    'Add StickLength Node (if applicable)
    If TypeName(cTank.StickLength) <> "Nothing" Then
        Set xParent = xDoc.createElement("StickLength")
        xTankNode.appendChild xParent
        For i% = (cTank.StickLength.Count - 1) To 0 Step -1
            'Add Node
            sNode = "k" & Trim$(CStr(i%))
            Set xNode = xDoc.createElement(sNode)
            xNode.Text = cTank.StickLength(Right$(sNode, 1)) 'turns "K0" into "0")
            xParent.appendChild xNode
        Next i%
    End If
    
    'Add ContentCG Node
    If TypeName(cTank.ContentCG) <> "Nothing" Then
        Set xParent = xDoc.createElement("ContentCG")
        xTankNode.appendChild xParent
        For i% = (cTank.ContentCG.Count - 1) To 0 Step -1
            'Add Node
            sNode = "k" & Trim$(CStr(i%))
            Set xNode = xDoc.createElement(sNode)
            xNode.Text = cTank.ContentCG(Right$(sNode, 1)) 'turns "K0" into "0")
            xParent.appendChild xNode
        Next i%
    End If
    
    'Add ContentVCG Node
    If TypeName(cTank.ContentVCG) <> "Nothing" Then
        Set xParent = xDoc.createElement("ContentVCG")
        xTankNode.appendChild xParent
        For i% = (cTank.ContentVCG.Count - 1) To 0 Step -1
            'Add Node
            sNode = "k" & Trim$(CStr(i%))
            Set xNode = xDoc.createElement(sNode)
            xNode.Text = cTank.ContentVCG(Right$(sNode, 1)) 'turns "K0" into "0")
            xParent.appendChild xNode
        Next i%
    End If
    
    FillTankNode = "" 'No error if we reached this point

CleanUp:
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillTankNode(xDoc,xTankNode,cTank)", Array(xDoc, xTankNode, cTank)
End Function
    

Public Function FillChassisNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef cTruck As clsTruck) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim sErrMsg As String
    Dim i%
    Dim sNodeName() As String
    Dim sNodes As String
    On Error GoTo errHandler
    
    FillChassisNode = "Unknown Error in modXML:FillChassisNode"
    'Create Chassis node
    Set xParent = xDoc.createElement("Chassis")
    Set xNode = xDoc.selectSingleNode("Truck")
    If TypeName(xNode) = "Nothing" Then
        'Not a Truck file, so parent node is 'ALObject'
        xDoc.selectSingleNode("ALObject").appendChild xParent
    Else
        'Truck file, so parent node is 'Truck'
        xDoc.selectSingleNode("Truck").appendChild xParent
    End If
        
    'Add stand-alone Chassis Nodes
    sNodes = "FullName|DisplayName|WB|TwinSteerSeparation|BackOfCab|WtFront|WtRear|WtLimitFront|WtLimitRear|WtLimitTotal|TandemSpacing|PlacementAllowable|Placement|StreetSideStd|CurbSideStd|WheelDia|WheelY|FrameRailHt"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        xNode.Text = CallByNameEx(cTruck.Chassis, sNodeName(i%), VbGet)
        If sNodeName(i%) = "PlacementAllowable" Then xNode.Text = cTruck.Chassis.PlacementAllowableString
        If sNodeName(i%) = "Placement" Then xNode.Text = cTruck.Chassis.PlacementString
        If xNode.Text <> "" Then
            xParent.appendChild xNode
        End If
    Next i%
    
    If cTruck.Chassis.Tags.Count <> 0 Then
        'Add Tags sub-node
        Set xNode = xDoc.createElement("Tags")
        xParent.appendChild xNode
        'Make 'Tags' node the parent
        Set xParent = xNode
        'Add 'Tag' nodes
        For i% = 1 To cTruck.Chassis.Tags.Count
            Set xNode = xDoc.createElement("Tag")
            xParent.appendChild xNode
            sErrMsg = FillTagNode(xDoc, xNode, cTruck.Chassis.Tags(i%))
        Next i%
    End If
    
    FillChassisNode = "" 'No error if we reached this point
    
CleanUp:
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillChassisNode(xDoc,cTruck)", Array(xDoc, cTruck)
End Function

Private Function FillTagNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef xTagNode As IXMLDOMNode, cTag As clsTag) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim sNodes As String
    Dim sNodeName() As String
    Dim i%
    On Error GoTo errHandler
    
    FillTagNode = "Unknown Error in modXML:FillTagNode"
    'Add stand-alone Tag Nodes
    sNodes = "Location|Weight|WtLimit|ForceToPressure|DownwardForce"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        xNode.Text = CallByNameEx(cTag, sNodeName(i%), VbGet)
        xTagNode.appendChild xNode
    Next i%
    
    FillTagNode = "" 'No error if we reached this point
    
CleanUp:
    Set xNode = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillTagNode(xDoc,xTagNode,cTag)", Array(xDoc, xTagNode, cTag)
End Function


Public Function FillComponentsNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef cTruck As clsTruck, Optional colComponents As Collection = Nothing) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim sErrMsg As String
    Dim i%
    Dim x%
    On Error GoTo errHandler
    
    FillComponentsNode = "Unknown Error in modXML:FillComponentsNode"
    'Create Components node as 'Parent'
    Set xParent = xDoc.createElement("Components")
    Set xNode = xDoc.selectSingleNode("Truck")
    If TypeName(xNode) = "Nothing" Then
        'Not a Truck file, so parent node is 'ALObject'
        xDoc.selectSingleNode("ALObject").appendChild xParent
    Else
        'Truck file, so parent node is 'Truck'
        xDoc.selectSingleNode("Truck").appendChild xParent
    End If
        
    'Add 'Component' nodes
    If TypeName(colComponents) = "Nothing" Then
        'Save all components
        For i% = 1 To cTruck.Components.Count
            Set xNode = xDoc.createElement("Component")
            xParent.appendChild xNode
            sErrMsg = FillComponentNode(xDoc, xNode, cTruck.Components(i%))
        Next i%
    Else
        'Only save selected components
        For i% = 1 To cTruck.Components.Count
            For x% = 1 To colComponents.Count
                If i% = CInt(colComponents.Item(x%)) Then
                    'This component should be included
                    Set xNode = xDoc.createElement("Component")
                    xParent.appendChild xNode
                    sErrMsg = FillComponentNode(xDoc, xNode, cTruck.Components(i%))
                    Exit For
                End If
            Next x%
        Next i%
    End If
    
    FillComponentsNode = "" 'No error if we reached this point
    
CleanUp:
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillComponentsNode(xDoc,cTruck)", Array(xDoc, cTruck)
End Function


Public Function FillComponentNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef xComponentNode As IXMLDOMNode, cComponent As clsComponent) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim sNodes As String
    Dim sNodeName() As String
    Dim sErrMsg As String
    Dim i%
    On Error GoTo errHandler
    
    
    'Add stand-alone Component Nodes
    sNodes = "LocationReference|FullName|DisplayName|Offset|EmptyWeight|EmptyCG|ContentsType|InstallationNotes|PlacementAllowable|Placement|StreetSideStd|CurbSideStd"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        xNode.Text = CallByNameEx(cComponent, sNodeName(i%), VbGet)
        If sNodeName(i%) = "LocationReference" Then xNode.Text = cComponent.LocationReferenceString
        If sNodeName(i%) = "ContentsType" Then xNode.Text = cComponent.ContentsTypeString
        If sNodeName(i%) = "PlacementAllowable" Then xNode.Text = cComponent.PlacementAllowableString
        If sNodeName(i%) = "Placement" Then xNode.Text = cComponent.PlacementString
        xComponentNode.appendChild xNode
    Next i%
    
    If cComponent.FillRelationShips.Count > 0 Then
        'Add FillRelationships sub-nodes
        Set xNode = xDoc.createElement("FillRelationships")
        xComponentNode.appendChild xNode
        'Make 'FillRelationships' node the parent
        Set xParent = xNode
        'Add 'FillRelationship' nodes
        For i% = 1 To cComponent.FillRelationShips.Count
            Set xNode = xDoc.createElement("FillRelationship")
            xParent.appendChild xNode
            sErrMsg = FillFillRelationshipNode(xDoc, xNode, cComponent.FillRelationShips(i%))
        Next i%
    End If
    
    If cComponent.ContentsType <> ctNone Then
        'Add Capacity sub-node
        Set xNode = xDoc.createElement("Capacity")
        xComponentNode.appendChild xNode
        sErrMsg = FillCapacityNode(xDoc, xNode, cComponent.Capacity)
    End If
        
    FillComponentNode = "" 'No error if we reached this point
    
CleanUp:
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillComponentNode(xDoc,xComponentNode,cComponent)", Array(xDoc, xComponentNode, _
         cComponent)
End Function

Private Function FillFillRelationshipNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef xFillRelationshipNode As IXMLDOMNode, cFillRelationship As clsFillRelationship) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim sNodes As String
    Dim sNodeName() As String
    Dim i%
    On Error GoTo errHandler
    
    FillFillRelationshipNode = "Unknown Error in modXML:FillFillRelationshipNode"
    'Add stand-alone FillRelationship Nodes
    sNodes = "ParentProduct|Multiplier|Offset"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        xNode.Text = CallByNameEx(cFillRelationship, sNodeName(i%), VbGet)
        If sNodeName(i%) = "ParentProduct" Then xNode.Text = cFillRelationship.ParentProductString
        xFillRelationshipNode.appendChild xNode
    Next i%
    FillFillRelationshipNode = "" 'No error if we reached this point
    
CleanUp:
    Set xNode = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillFillRelationshipNode(xDoc,xFillRelationshipNode,cFillRelationship)", _
         Array(xDoc, xFillRelationshipNode, cFillRelationship)
End Function


Private Function FillCapacityNode(ByRef xDoc As MSXML2.DOMDocument30, ByRef xCapacityNode As IXMLDOMNode, cCapacity As clsCapacity) As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim sNode As String
    Dim sNodes As String
    Dim sNodeName() As String
    Dim i%
    On Error GoTo errHandler
    
    FillCapacityNode = "Unknown Error in modXML:FillCapacityNode"
    'Add stand-alone Capacity Nodes
    sNodes = "DefaultWtContents|DefaultVolContents|DensityContents|UsesSightGauge|Volume"
    sNodeName = Split(sNodes, "|")
    For i% = 0 To UBound(sNodeName)
        'add node
        Set xNode = xDoc.createElement(sNodeName(i%))
        xNode.Text = CallByNameEx(cCapacity, sNodeName(i%), VbGet)
        If xNode.Text <> "" Then
            xCapacityNode.appendChild xNode
        End If
    Next i%
    
    'Add StickLength Node (if applicable)
    If TypeName(cCapacity.StickLength) <> "Nothing" Then 'avoids potential error
        If cCapacity.StickLength.Count > 0 Then 'ONLY save if count greater than zero
            Set xParent = xDoc.createElement("StickLength")
            xCapacityNode.appendChild xParent
            For i% = (cCapacity.StickLength.Count - 1) To 0 Step -1
                'Add Node
                sNode = "k" & Trim$(CStr(i%))
                Set xNode = xDoc.createElement(sNode)
                xNode.Text = cCapacity.StickLength(Right$(sNode, 1)) 'turns "K0" into "0")
                xParent.appendChild xNode
            Next i%
        End If
    End If
    
    'Add ContentCG Node (if applicable)
    If TypeName(cCapacity.ContentCG) <> "Nothing" Then 'avoids potential error
        If cCapacity.ContentCG.Count > 0 Then 'ONLY save if count greater than zero
            Set xParent = xDoc.createElement("ContentCG")
            xCapacityNode.appendChild xParent
            For i% = (cCapacity.ContentCG.Count - 1) To 0 Step -1
                'Add Node
                sNode = "k" & Trim$(CStr(i%))
                Set xNode = xDoc.createElement(sNode)
                xNode.Text = cCapacity.ContentCG(Right$(sNode, 1)) 'turns "K0" into "0")
                xParent.appendChild xNode
            Next i%
        End If
    End If
    FillCapacityNode = "" 'No error if we reached this point
CleanUp:
    Set xNode = Nothing
    Set xParent = Nothing
    Exit Function
errHandler:
    ErrorIn "modXML.FillCapacityNode(xDoc,xCapacityNode,cCapacity)", Array(xDoc, xCapacityNode, _
         cCapacity)
End Function


