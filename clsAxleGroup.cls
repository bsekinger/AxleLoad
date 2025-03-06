VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAxleGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public sGroup As String 'e.g.  1-2, 1-2-3, 2-3, etc.

Public sDescription As String 'e.g.  "Front-Tag-Rear", etc.

Public Spacing As Long 'extreme axle spacing of group **(ft)**

Public AllowableLd As Double 'calculated allowable load for group (kg)

Public ActualLd As Double 'sum axle-loadings of group (kg)

Public AxleLoc As Double 'ONLY applicable to singe Axle

Public NumAxles As Long  'ONLY applicable to axle groups

Public MinLoading As Double 'zero for tags

Public ForceToPressure As Double 'only valid for tag axles

