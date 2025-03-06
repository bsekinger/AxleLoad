VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Location As Double 'Distance from front axle (m)

Public Weight As Double 'weight of equipment (kg)

Public WtLimit As Double 'Max allowable Downward Force (kg)

Public ForceToPressure As Double 'mult. 'DownwardForce' by this value to get air pressure (no units)

'--- Properties used by Solver ---------------------

Public DownwardForce As Double 'positive value (kg)
