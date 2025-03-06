VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTruckInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Truck Information"
   ClientHeight    =   6360
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   8160
   Icon            =   "frmTruckInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   6720
      TabIndex        =   1
      Top             =   5880
      Width           =   1332
   End
   Begin MSComctlLib.TreeView treTruck 
      Height          =   1812
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   7932
      _ExtentX        =   13991
      _ExtentY        =   3196
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox txtInfo 
      Height          =   3072
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   7932
   End
   Begin VB.Label lblLabel 
      Caption         =   "Capacity data and other general information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6132
   End
   Begin VB.Label lblLabel 
      Caption         =   "XML Source   (all values are in SI units - kg, m, liter)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   5892
   End
End
Attribute VB_Name = "frmTruckInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Const COMMENT_COLOR = &H8000&
Private Const ATTRIBUTE_COLOR = &H808080
Private Const ELEMENT_COLOR = &H0
Private Const VALUE_COLOR = &HFF0000
Private Const PROCESSING_INSTRUCTION_COLOR = &H80FF

Private cCurTruck As clsTruck 'used in this form as a link to the active truck

Public Sub Edit(strTruckFile As String, cTruck As clsTruck)
    Dim xDoc As DOMDocument
    
    Set cCurTruck = cTruck
    
    ' Load the XML file into the DOMDocument.
    Set xDoc = New DOMDocument
    xDoc.Load strTruckFile
    
    LoadTreeFromXML xDoc, treTruck, True
    UpdateInfoBox
    
    Me.Show

End Sub


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cCurTruck = Nothing
End Sub


Public Sub LoadTreeFromXML(objDOM As MSXML2.DOMDocument30, tvtreeview As TreeView, blnAttribNodes As Boolean)
    Dim nod As MSXML2.IXMLDOMNode
    Dim nodx As Node
    Dim f$
    Dim lLoc As Long
    
    tvtreeview.Nodes.Clear
    f$ = Replace(objDOM.url, "%20", " ")
    lLoc = InStrRev(f$, "/")
    f$ = Mid$(f$, lLoc + 1)
    Set nodx = tvtreeview.Nodes.Add(, , , f$)
    nodx.Expanded = True
    For Each nod In objDOM.childNodes
        DisplayNode nod, nodx, tvtreeview, blnAttribNodes
    Next
      
    Set nod = Nothing
    Set nodx = Nothing
End Sub

Private Sub DisplayNode(objNode As MSXML2.IXMLDOMNode, objtvParentNode As Node, _
                tvtreeview As TreeView, blnAttribNodes As Boolean)
                
    Dim nod As IXMLDOMNode
    Dim atrib As IXMLDOMAttribute
    Dim nodx As Node
    Dim noda As Node
    
    Select Case objNode.nodeType
    
    Case NODE_ELEMENT
        Set nodx = tvtreeview.Nodes.Add(objtvParentNode, tvwChild, , _
                    objNode.baseName)
        nodx.ForeColor = ELEMENT_COLOR
        nodx.Expanded = True
    Case NODE_COMMENT
        Set nodx = tvtreeview.Nodes.Add(objtvParentNode, tvwChild, , _
                    "* " & objNode.Text & " *")
        nodx.ForeColor = COMMENT_COLOR
        nodx.Expanded = True
    Case NODE_TEXT
        Set nodx = tvtreeview.Nodes.Add(objtvParentNode, tvwChild, , _
                    objNode.Text)
        nodx.ForeColor = VALUE_COLOR
        nodx.Expanded = True
    Case Else
        Set nodx = objtvParentNode
    End Select
                   
    If Not objNode.Attributes Is Nothing Then
        For Each atrib In objNode.Attributes
            Select Case atrib.nodeType
            Case NODE_ATTRIBUTE
                If blnAttribNodes Then
                    Set noda = tvtreeview.Nodes.Add(nodx, tvwChild, , _
                        "[" & atrib.baseName & "=" & atrib.Value & "]")
                    noda.ForeColor = ATTRIBUTE_COLOR
                Else
                    nodx.Text = nodx.Text & _
                        " [" & atrib.baseName & "=" & atrib.Value & "]"
                End If
                
            Case NODE_PROCESSING_INSTRUCTION
                Set noda = tvtreeview.Nodes.Add(nodx, tvwChild, , _
                        "{{" & atrib.baseName & "=" & atrib.Value & "}}")
                noda.ForeColor = PROCESSING_INSTRUCTION_COLOR
            End Select
        Next
    End If
    
    If objNode.childNodes.length > 0 Then
        For Each nod In objNode.childNodes
            DisplayNode nod, nodx, tvtreeview, blnAttribNodes
        Next
    End If
End Sub


Private Function UpdateInfoBox() As Boolean
    'This sub fills the Info box with all sorts of loading information
    Dim a$, b$
    Dim i%
    Dim dblLoad As Double
    Dim dblTemp As Double
    Dim cEmulTanks As New clsTanks
    Dim cANTanks As New clsTanks
    
    
    Set cANTanks = ANTanks(cCurTruck)
    Set cEmulTanks = EmulsionTanks(cCurTruck)
        
    'Truck info
    txtInfo.SelText = "Unit: " & cCurTruck.Description & vbCrLf
    txtInfo.SelText = "Tread SN" & cCurTruck.SN & vbCrLf
    txtInfo.SelText = "Owner: " & cCurTruck.Owner & vbCrLf
    txtInfo.SelText = "Std Mount Body: " & cCurTruck.IsStandardMount & vbCrLf
    txtInfo.SelText = vbCrLf
    
    'Chassis Info
    txtInfo.SelText = "Chassis: " & cCurTruck.Chassis.DisplayName & vbCrLf
    dblTemp = cCurTruck.Chassis.WB * cGlobalInfo.DistanceUnits.Multiplier
    a$ = Format(dblTemp, "# ") & cGlobalInfo.DistanceUnits.Display
    txtInfo.SelText = "Wheel Base: " & a$ & vbCrLf
    'Manufacturer's Max Load Spec for Front axle/tire
    dblLoad = cCurTruck.Chassis.WtLimitFront * cGlobalInfo.MassUnits.Multiplier
    a$ = "MfgLimit Front: "
    b$ = Format(dblLoad, "# ") & cGlobalInfo.MassUnits.Display
    txtInfo.SelText = a$ & b$ & vbCrLf
    'Manufacturer's Max Load Spec for Rear axle/tire
    dblLoad = cCurTruck.Chassis.WtLimitRear * cGlobalInfo.MassUnits.Multiplier
    a$ = "MfgLimit Rear: "
    b$ = Format(dblLoad, "# ") & cGlobalInfo.MassUnits.Display
    txtInfo.SelText = a$ & b$ & vbCrLf
    'Tag Specs
    For i% = 1 To cCurTruck.Chassis.Tags.Count
        txtInfo.SelText = "Tag #" & i% & vbCrLf
        'Mfg Limit
        a$ = "   MfgLimit: "
        dblTemp = cCurTruck.Chassis.Tags(i%).WtLimit * cGlobalInfo.MassUnits.Multiplier
        b$ = Format(dblTemp, "# ") & cGlobalInfo.MassUnits.Display
        txtInfo.SelText = a$ & b$ & vbCrLf
        'Location
        a$ = "   Distance from front axle: "
        dblTemp = cCurTruck.Chassis.Tags(i%).Location * cGlobalInfo.DistanceUnits.Multiplier
        b$ = Format(dblTemp, "#.## ") & cGlobalInfo.DistanceUnits.Display
        txtInfo.SelText = a$ & b$ & vbCrLf
    Next i%
    
    'Tank Capacities
    txtInfo.SelText = vbCrLf
    For i% = 1 To cCurTruck.Body.Tanks.Count
        txtInfo.SelText = cCurTruck.Body.Tanks(i%).DisplayName & vbCrLf
        'Volume Capacity
        dblTemp = cCurTruck.Body.Tanks(i%).Volume * cGlobalInfo.VolumeUnits.Multiplier
        a$ = Format(dblTemp, "# ") & cGlobalInfo.VolumeUnits.Display
        txtInfo.SelText = "   Volume Capacity: " & a$ & vbCrLf
        Select Case cCurTruck.Body.Tanks(i%).TankType
        Case ttAN
            dblTemp = cCurTruck.Body.Tanks(i%).Volume * cGlobalInfo.DensityAN * cGlobalInfo.MassUnits.Multiplier
            a$ = Format(dblTemp, "# ") & cGlobalInfo.MassUnits.Display
            a$ = "   AN Capacity: " & a$
            b$ = " (S.G. = " & Format(cGlobalInfo.DensityAN, "0.##") & ")"
            txtInfo.SelText = a$ & b$ & vbCrLf
        Case ttEmulsion
            dblTemp = cCurTruck.Body.Tanks(i%).Volume * cGlobalInfo.DensityEmul * cGlobalInfo.MassUnits.Multiplier
            a$ = Format(dblTemp, "# ") & cGlobalInfo.MassUnits.Display
            a$ = "   Emulsion Capacity: " & a$
            b$ = " (S.G. = " & Format(cGlobalInfo.DensityEmul, "0.##") & ")"
            txtInfo.SelText = a$ & b$ & vbCrLf
        Case ttDual
            'AN
            dblTemp = cCurTruck.Body.Tanks(i%).Volume * cGlobalInfo.DensityAN * cGlobalInfo.MassUnits.Multiplier
            a$ = Format(dblTemp, "# ") & cGlobalInfo.MassUnits.Display
            a$ = "   AN Capacity: " & a$
            b$ = " (S.G. = " & Format(cGlobalInfo.DensityAN, "0.##") & ")"
            txtInfo.SelText = a$ & b$ & vbCrLf
            'Emulsion
            dblTemp = cCurTruck.Body.Tanks(i%).Volume * cGlobalInfo.DensityEmul * cGlobalInfo.MassUnits.Multiplier
            a$ = Format(dblTemp, "# ") & cGlobalInfo.MassUnits.Display
            a$ = "   Emulsion Capacity: " & a$
            b$ = " (S.G. = " & Format(cGlobalInfo.DensityEmul, "0.##") & ")"
            txtInfo.SelText = a$ & b$ & vbCrLf
        End Select
        txtInfo.SelText = vbCrLf
    Next
        
    'Capacities of components (water, gassing, etc.)
    For i% = 1 To cCurTruck.Components.Count
        Select Case cCurTruck.Components(i%).ContentsType
        Case ctNone
            'Don't show
            a$ = ""
            b$ = ""
        Case ctOther
            a$ = cCurTruck.Components(i%).DisplayName
            b$ = ""
            If cCurTruck.Components(i%).Capacity.Volume > 0 Then
                dblTemp = cCurTruck.Components(i%).Capacity.Volume * cGlobalInfo.VolumeUnits.Multiplier
                b$ = ": " & Format(dblTemp, "# ") & cGlobalInfo.VolumeUnits.Display
            End If
        Case Else
            a$ = cCurTruck.Components(i%).DisplayName & ": "
            dblTemp = cCurTruck.Components(i%).Capacity.Volume * cGlobalInfo.VolumeUnits.Multiplier
            b$ = Format(dblTemp, "# ") & cGlobalInfo.VolumeUnits.Display
        End Select
        txtInfo.SelText = a$ & b$ & vbCrLf
    Next
    Set cANTanks = Nothing
    Set cEmulTanks = Nothing
End Function

