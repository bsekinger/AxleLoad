VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapacity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public DefaultWtContents As String '"" = don't adjust by weight

Public DefaultVolContents As String '"" = don't adjust by volume

Public DensityContents As Double 'set by program if contents are of known type

Public UsesSightGauge As Boolean

Public Volume As Double

Public StickLength As Collection
' K0 >> Item 1
' K1 >> Item 2
'...
' K5 >> Item 6

Public ContentCG As Collection
' K0 >> Item 1
' K1 >> Item 2
'...
' K5 >> Item 6

'--- Properties used by Solver ---------------------
Public MaxHt As Double 'Calculated using Volume and StickLength

Public CurStkHt As Double 'set by program

Public CurVol As Double 'set by program

Public CurContentCG As Double 'set by program
