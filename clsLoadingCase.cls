VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadingCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public VCOG As Double 'Vertical COG of Product (AN & Emul)

Public TotalProductMass As Double 'Sum of Product (AN & Emul) in kg

Private m_TankConfig As Collection 'either AN (0) or Emulsion (1)

Private m_ContentVol As Collection 'L


Public Sub SetTankConfig(ByVal Value As Double, ByVal sIndex As String)
    'Set value for a member of the collection
    m_TankConfig.Remove (sIndex)
    m_TankConfig.Add Value, sIndex
End Sub

Public Property Get TankConfig(ByVal sIndex As String) As Variant
        TankConfig = m_TankConfig.Item(sIndex)
End Property


Public Sub SetContentVol(ByVal Value As Double, ByVal sIndex As String)
    'Set value for a member of the collection
    m_ContentVol.Remove (sIndex)
    m_ContentVol.Add Value, sIndex
End Sub


Public Property Get ContentVol(ByVal sIndex As String) As Variant
        ContentVol = m_ContentVol.Item(sIndex)
End Property


Public Sub InitTanks(cTruck As clsTruck)
    Dim i%
    
    Set m_TankConfig = New Collection
    Set m_ContentVol = New Collection
    For i% = 1 To cTruck.Body.Tanks.Count
        m_TankConfig.Add cTruck.Body.Tanks(i%).CurTankUse, CStr(i%)
        m_ContentVol.Add cTruck.Body.Tanks(i%).CurVol, CStr(i%)
    Next i%
    'Create worst-case moment to start with
    VCOG = 1.79E+308 'biggest number possible
End Sub

