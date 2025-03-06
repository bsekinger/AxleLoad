VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum TankUse
    ttAN = 0
    ttEmulsion = 1
    ttDual = 2
End Enum

Public DisplayName As String

Public TankType As TankUse

Public Volume As Double

Public MaxMass As Double 'Max allowable mass in tank (for On-Road use)

Public MaxMassDesc As String 'Reason for mass limit (e.g. "DOT 406")

Public StickLength As Collection
' K0 >> Item 1
' K1 >> Item 2
'...
' K5 >> Item 6

Public ContentCG As Collection 'COG, front to back
' K0 >> Item 1
' K1 >> Item 2
'...
' K5 >> Item 6

Public ContentVCG As Collection 'COG, vertical
' K0 >> Item 1
' K1 >> Item 2
'...
' K5 >> Item 6

Public CurTankUse As TankUse 'either AN or Emulsion

'--- Properties used by Solver ---------------------
Public MaxHt As Double 'Calculated using Volume and StickLength

Public CurStkHt As Double 'set by program

Public CurVol As Double 'set by program

Public DensityContents As Double 'set by program for either AN or Emul

Public CurContentCG As Double 'set by program

Public CurContentVCG As Double 'set by program


'--- Calculated Properties  ---------------------
Public Property Get TankTypeString() As String
    'Read-only string representation of Enum for writing to file
    If TankType = ttAN Then
        TankTypeString = "AN"
    ElseIf TankType = ttEmulsion Then
        TankTypeString = "Emulsion"
    Else
        TankTypeString = "Dual"
    End If
End Property

Public Property Get CurTankUseString() As String
    'Read-only string representation of Enum for writing to file
    If CurTankUse = ttEmulsion Then
        CurTankUseString = "Emulsion"
    Else
        CurTankUseString = "AN"
    End If
End Property

