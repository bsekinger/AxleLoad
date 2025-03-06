VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public DensityAN As Double

Public DensityEmul As Double

Public DensityFuel As Double

Public DensityWater As Double

Public DensityGasA As Double

Public DensityGasB As Double

Public DensityAdditive As Double

Public TruckFolder As String

Public ALObjectFolder As String

Public AvoidLowFill As Boolean

Public VolumeUnits As clsUnits

Public DistanceUnits As clsUnits

Public MassUnits As clsUnits

'--- Non-persistent properties used in software  ---------------------
Public FillMethod_Fuel As Integer '0=Auto, 1=Empty, 2=Full
Public FillMethod_Water As Integer '0=Auto, 1=Empty, 2=Full
Public FillMethod_GasA As Integer '0=Auto, 1=Empty, 2=Full
Public FillMethod_GasB As Integer '0=Auto, 1=Empty, 2=Full
Public FillMethod_Additive As Integer '0=Auto, 1=Empty, 2=Full


