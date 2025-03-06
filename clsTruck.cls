VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public SN As String

Public Owner As String

Public ShipDate As Date

Public BridgeLaw As String

Public Description As String

Public IsStandardMount As Boolean

Public BodyLocation As Double

Public WtAdjustFront As Double

Public WtAdjustRear As Double

Public Body As clsBody

Public Chassis As clsChassis

Public Components As clsComponents

Public CreateVersion As clsVersion

