VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum ALObjectType
    xptBody = 0
    xptChassis = 1
    xptComponent = 2
End Enum

Public DisplayName As String 'What's displayed in list-box

Public FullName As String 'Verbose description

Public File As String 'Filename w/ complete path

Public ObjectType As ALObjectType

Public Index As Long 'Used to link to a listbox.itemdata

