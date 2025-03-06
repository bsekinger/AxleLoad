VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Public Major As Integer

Public Minor As Integer

Public Revision As Integer




'local variable(s) to hold property value(s)
Private mvarTextVal As String 'local copy


Public Property Get TextVal() As String
Attribute TextVal.VB_UserMemId = 0
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.TextVal
    mvarTextVal = Trim$(CStr(Major)) & "." & _
                  Trim$(CStr(Minor)) & "." & _
                  Trim$(CStr(Revision))

    TextVal = mvarTextVal
End Property

Public Property Let TextVal(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.TextVal = "5.1.2"
    Dim i%
    Dim intIndent As Integer
    Dim a$
    
    mvarTextVal = Trim$(vData)
    mvarTextVal = Replace(mvarTextVal, Chr$(34), "") 'get rid of quotation marks
    mvarTextVal = Replace(mvarTextVal, "'", "") 'get rid of single-quotation marks
    
    intIndent = 0
    
    Major = 0
    Minor = 0
    Revision = 0
    
    For i% = 1 To Len(mvarTextVal)
        c$ = Mid$(mvarTextVal, i%, 1)
        If c$ = "." Or i% = Len(mvarTextVal) Then
            If a$ = "" Then a$ = c$
            intIndent = intIndent + 1
            Select Case intIndent
            Case 1
                Major = CInt(a$)
            Case 2
                Minor = CInt(a$)
            Case 3
                Revision = CInt(a$)
            End Select
            a$ = ""
        Else
            a$ = a$ & c$
        End If
    Next
End Property




