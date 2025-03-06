Attribute VB_Name = "modSystem"
Option Explicit
Declare Function CoCreateGuid_Alt Lib "OLE32.DLL" Alias "CoCreateGuid" (pGuid As Any) As Long
Declare Function StringFromGUID2_Alt Lib "OLE32.DLL" Alias "StringFromGUID2" (pGuid As Any, ByVal address As Long, ByVal Max As Long) As Long

'----------------------------------------------------------------
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'----------------------------------------------------------------


'Declares for BrowseFolders-----------------------
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
'-------------------------------------------------

'Declares for LV_AutoSizeColumn---------------------------------
Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
'----------------------------------------------------------------

'--- API function to read a key state
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
'e.g. -   Capslock = CBool(GetKeyState(vbKeyCapital) And 1)
'----------------------------------------------------------------

'--- API declarations for GetSystemMetrics----------------
Declare Function GetSystemMetrics Lib "user32" (ByVal nindex As Long) As Long
'---------------------------------------------------------------------

'--- API function Initialize Commom Controls
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200
'----------------------------------------------------------------

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Function CallByNameEx(Obj As Object, _
    ProcName As String, CallType As VbCallType, _
    Optional vArgsArray As Variant)
    Dim oTLI As Object
    Dim ProcID As Long
    Dim numArgs As Long
    Dim i As Long
    Dim V()
    
    On Error GoTo Handler
    
    Set oTLI = CreateObject("TLI.TLIApplication")
    ProcID = oTLI.InvokeID(Obj, ProcName)
    
    If IsMissing(vArgsArray) Then
        CallByNameEx = oTLI.InvokeHook( _
            Obj, ProcID, CallType)
    End If
    
    If IsArray(vArgsArray) Then
        numArgs = UBound(vArgsArray)
        ReDim V(numArgs)
        For i = 0 To numArgs
            V(i) = vArgsArray(numArgs - i)
        Next i
        CallByNameEx = oTLI.InvokeHookArray( _
            Obj, ProcID, CallType, V)
    End If
Exit Function

Handler:
    Debug.Print Err.Number, Err.Description
End Function

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
 Dim c As ColumnHeader
 
    If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
            SendMessage LV.hwnd, LVM_FIRST + 30, c.Index - 1, -1
        Next
    Else
        SendMessage LV.hwnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
End Sub


'Create a GUID
Function CreateGUID() As String
    Dim res As String, resLen As Long, GUID(15) As Byte
    res = Space$(128)
    CoCreateGuid_Alt GUID(0)
    resLen = StringFromGUID2_Alt(GUID(0), ByVal StrPtr(res), 128)
    CreateGUID = Left$(res, resLen - 1)
End Function

Public Sub SetGridColumnWidth(grd As MSFlexGrid, Optional bExtendLast As Boolean = False)
    'params:    ms flexgrid control
    'purpose:   sets the column widths to the
    '           lengths of the longest string in the column
    'requirements:  the grid must have the same
    '               font as the underlying form

    Dim InnerLoopCount As Long
    Dim OuterLoopCount As Long
    Dim lngLongestLen As Long
    Dim sLongestString As String
    Dim lngColWidth As Long
    Dim szCellText As String
    Dim lTotalWidth As Long

    For OuterLoopCount = 0 To grd.Cols - 1
        sLongestString = ""
        lngLongestLen = 0

        'grd.Col = OuterLoopCount
        For InnerLoopCount = 0 To grd.Rows - 1
            szCellText = grd.TextMatrix(InnerLoopCount, OuterLoopCount)
            'grd.Row = InnerLoopCount
            'szCellText = Trim$(grd.Text)
            If Len(szCellText) > lngLongestLen Then
                lngLongestLen = Len(szCellText)
                sLongestString = szCellText
            End If
        Next
        lngColWidth = grd.Parent.TextWidth(sLongestString & "H.")

        'add 100 for more readable spreadsheet
        grd.ColWidth(OuterLoopCount) = lngColWidth + 200
        lTotalWidth = lTotalWidth + grd.ColWidth(OuterLoopCount) + grd.GridLineWidth
    Next
    
    If bExtendLast Then
        lTotalWidth = lTotalWidth + 10 * grd.GridLineWidth _
        + GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX  'twips
                     
        If lTotalWidth < grd.Width Then
            grd.ColWidth(grd.Cols - 1) = (grd.Width - lTotalWidth) + grd.ColWidth(grd.Cols - 1)
        End If
    End If
End Sub

Public Function BrowseFolders(ByVal hwnd As Long, Optional ByVal DialogTitle As String = "@@@") As String
    'show the dialog to select a folder
    Dim BI As BROWSEINFO
    Dim lID As Long
    Dim szPath As String
    
    BI.hOwner = hwnd
    If DialogTitle <> "@@@" Then BI.lpszTitle = DialogTitle
    'return physical folders only
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    szPath = Space$(1024)
    lID = SHBrowseForFolder(BI)
    If SHGetPathFromIDList(ByVal lID, ByVal szPath) Then
        BrowseFolders = Left$(szPath, InStr(szPath, vbNullChar) - 1)
    Else
        BrowseFolders = ""
    End If
End Function


