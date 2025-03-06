VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChassis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public FullName As String

Public DisplayName As String

Public WB As Double

Public TwinSteerSeparation As Double '(0=Single Front axle) Dist. b/w twin-steer axles

Public BackOfCab As Double

Public WtFront As Double

Public WtRear As Double

Public WtLimitFront As Double

Public WtLimitRear As Double

Public WtLimitTotal As Double

Public TandemSpacing As Double

Public PlacementAllowable As PlacementAllowableLocation 'Either Side, StreetSideStd, CurbSideStd, or Center

Public Placement As PlacementLocation 'Not Placed, StreetSideStd, CurbSideStd, or Center

Public StreetSideStd  As String 'name of file (dxf) to show for Steet-Side Std-Mt truck

Public CurbSideStd  As String 'name of file (dxf) to show for Curb-Side Std-Mt truck

Public WheelDia As Double

Public WheelY As Double

Public FrameRailHt As Double

Public Tags As clsTags

'--- Non-persistent properties used in software  ---------------------
Public GUID As String 'assigned whenever a chassis is read from an XML (links imported components)

'--- Calculated Properties  ---------------------
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


