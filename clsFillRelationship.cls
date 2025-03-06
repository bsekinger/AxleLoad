VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFillRelationship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'ComponentVol = [Parent1Vol * Multiplier1 + Offset1] + [Parent2Vol * Multiplier2 + Offset2]

Public Enum ParentProductType
    ptAN = 0
    ptEmulsion = 1
End Enum

Public ParentProduct As ParentProductType

Public Multiplier As Double

Public Offset As Double

'--- Calculated Properties  ---------------------
Public Property Get ParentProductString() As String
    Select Case ParentProduct
    Case ptAN
        ParentProductString = "AN"
    Case ptEmulsion
        ParentProductString = "Emulsion"
    End Select
End Property

