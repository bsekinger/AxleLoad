VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The private collection used to hold the real data
Private m_clsALObjects As Collection

Private Sub Class_Initialize()
    ' explicit assignment is slightly faster than auto-instancing
    Set m_clsALObjects = New Collection
End Sub

' Add a new clsALObject item to the collection

Public Sub Add(newItem As clsALObject, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
    ' TO DO: initialize new item's properties here
    ' ...
    ' add to the private collection
    m_clsALObjects.Add newItem, Key
End Sub

' Remove an item from the collection

Public Sub Remove(index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
    m_clsALObjects.Remove index
End Sub

' Return a clsALObject item from the collection

Public Function Item(index As Variant) As clsALObject
Attribute Item.VB_Description = "Returns a specific member of a Collection object either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_clsALObjects.Item(index)
End Function

' Return the number of items the collection

Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of members in a collection"
    Count = m_clsALObjects.Count
End Function

' Remove all the items from the collection

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all members from a Collection object"
    Set m_clsALObjects = New Collection
End Sub

' Implement support for enumeration (For Each)

Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_clsALObjects.[_NewEnum]
End Function


