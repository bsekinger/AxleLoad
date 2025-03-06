VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum ComponentContentsType
    ctNone = 0
    ctFuel = 1
    ctWater = 2
    ctGasA = 3
    ctGasB = 4
    ctAdditive = 5
    ctOther = 6
End Enum

'Because components could belong to the body object or the chassis object
Public Enum OriginReference
    orBodyOrigin = 0
    orFrontAxle = 1
End Enum


Public FullName As String

Public DisplayName As String

Public Offset As Double

Public LocationReference As OriginReference

Public EmptyWeight As Double

Public EmptyCG As Double

Public PlacementAllowable As PlacementAllowableLocation 'set to 'Either Side', StreetSideStd', 'CurbSideStd', or 'Center'

Public Placement As PlacementLocation 'Not Placed, StreetSideStd, CurbSideStd, or Center

Public StreetSideStd  As String 'name of file (dxf) to show for Steet-Side Std-Mt truck

Public CurbSideStd  As String 'name of file (dxf) to show for Curb-Side Std-Mt truck

Public ContentsType As ComponentContentsType

Public FillRelationShips As clsFillRelationships

Public Capacity As clsCapacity

Public InstallationNotes As String 'Details on where/how/when component is added

'--- Non-persistent properties used in software  ---------------------
Public ChassisGUID As String 'assigned when imported as part of a chassis file

'--- Calculated Properties  ---------------------
Public Property Get ContentsTypeString() As String
    'Read-only string representation of Enum for writing to file
    
    Select Case ContentsType
    Case ctNone
        ContentsTypeString = "None"
    Case ctFuel
        ContentsTypeString = "Fuel"
    Case ctWater
        ContentsTypeString = "Water"
    Case ctGasA
        ContentsTypeString = "GasA"
    Case ctGasB
        ContentsTypeString = "GasB"
    Case ctAdditive
        ContentsTypeString = "Additive"
    Case ctOther
        ContentsTypeString = "Other"
    End Select
End Property

Public Property Get LocationReferenceString() As String
    Select Case LocationReference
    Case orBodyOrigin
        LocationReferenceString = "Body"
    Case orFrontAxle
        LocationReferenceString = "Chassis"
    End Select
End Property

Public Property Get PlacementAllowableString() As String
    
    Select Case PlacementAllowable
    Case paEitherSide
        PlacementAllowableString = "Either Side"
    Case paStreetSideStd
        PlacementAllowableString = "StreetSideStd"
    Case paCurbSideStd
        PlacementAllowableString = "CurbSideStd"
    Case paCenter
        PlacementAllowableString = "Center"
    End Select
End Property


Public Property Get PlacementString() As String
    
    Select Case Placement
    Case plNotPlaced
        PlacementString = "Not Placed"
    Case plStreetSideStd
        PlacementString = "StreetSideStd"
    Case plCurbSideStd
        PlacementString = "CurbSideStd"
    Case plCenter
        PlacementString = "Center"
    End Select
End Property


