VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAxleGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' The private collection used to hold the real data
Private m_clsAxleGroups As Collection

Private Sub Class_Initialize()
    ' explicit assignment is slightly faster than auto-instancing
    Set m_clsAxleGroups = New Collection
End Sub

' Add a new clsAxleGroup item to the collection

Public Sub Add(newItem As clsAxleGroup, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
    ' TO DO: initialize new item's properties here
    ' ...
    ' add to the private collection
    m_clsAxleGroups.Add newItem, Key
End Sub

' Remove an item from the collection

Public Sub Remove(index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
    m_clsAxleGroups.Remove index
End Sub

' Return a clsAxleGroup item from the collection

Public Function Item(index As Variant) As clsAxleGroup
Attribute Item.VB_Description = "Returns a specific member of a Collection object either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_clsAxleGroups.Item(index)
End Function

' Return the number of items the collection

Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of members in a collection"
    Count = m_clsAxleGroups.Count
End Function

' Remove all the items from the collection

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all members from a Collection object"
    Set m_clsAxleGroups = New Collection
End Sub

' Implement support for enumeration (For Each)

Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_clsAxleGroups.[_NewEnum]
End Function


