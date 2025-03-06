VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DXFReaderClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'© Copyright 2000 Kadmos - Software Resolution
Private PI As Double

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Sub MoveDrawing(DXFReaderControl As DXFReader, DisplacementX As Single, DisplacementY As Single, DisplacementZ As Single)
    
    Dim Entity As DXFReaderEntity
    Dim Vertex As DXFReaderVertex
    
    With DXFReaderControl
        For Each Entity In .Entities
            If Entity.EntityType <> "IMAGE" Then
                If DisplacementX <> 0 Then
                    Entity.X0 = Entity.X0 + DisplacementX
                    Entity.X1 = Entity.X1 + DisplacementX
                    Entity.X2 = Entity.X2 + DisplacementX
                    Entity.X3 = Entity.X3 + DisplacementX
                End If
                If DisplacementY <> 0 Then
                    Entity.Y0 = Entity.Y0 + DisplacementY
                    Entity.Y1 = Entity.Y1 + DisplacementY
                    Entity.Y2 = Entity.Y2 + DisplacementY
                    Entity.Y3 = Entity.Y3 + DisplacementY
                End If
                If DisplacementZ <> 0 Then
                    Entity.Z0 = Entity.Z0 + DisplacementZ
                    Entity.Z1 = Entity.Z1 + DisplacementZ
                    Entity.z2 = Entity.z2 + DisplacementZ
                    Entity.z3 = Entity.z3 + DisplacementZ
                End If
                For Each Vertex In Entity.Vertexes
                    Vertex.X0 = Vertex.X0 + DisplacementX
                    Vertex.Y0 = Vertex.Y0 + DisplacementY
                    Vertex.Z0 = Vertex.Z0 + DisplacementZ
                Next Vertex
            Else
            
                If DisplacementX <> 0 Then
                    Entity.X0 = Entity.X0 + DisplacementX
                End If
                If DisplacementY <> 0 Then
                    Entity.Y0 = Entity.Y0 + DisplacementY
                End If
                If DisplacementZ <> 0 Then
                    Entity.Z0 = Entity.Z0 + DisplacementZ
                End If
            End If
            
        Next Entity
    End With
End Sub

Public Sub ScaleDrawing(DXFReaderControl As DXFReader, ScaleX As Single, ScaleY As Single, ScaleZ As Single)
    
    Dim Entity As DXFReaderEntity
    Dim Vertex As DXFReaderVertex
    Dim Block As DXFReaderBlock
    If ScaleX = 0 Then ScaleX = 1
    If ScaleY = 0 Then ScaleY = 1
    If ScaleZ = 0 Then ScaleZ = 1
    
    With DXFReaderControl
        For Each Entity In .Entities
            ScaleEntity Entity, ScaleX, ScaleY, ScaleZ
            
        Next Entity
        For Each Block In .Blocks
            If ScaleX <> 1 Then
                Block.X0 = ScaleX * Block.X0
            End If
            If ScaleY <> 1 Then
                Block.Y0 = ScaleY * Block.Y0
            End If
            
            If ScaleZ <> 1 Then
                Block.Z0 = ScaleY * Block.Z0
            End If
            For Each Entity In Block.Entities
                ScaleEntity Entity, ScaleX, ScaleY, ScaleZ
            Next Entity
        Next Block
        .LIMMINx = .LIMMINx * ScaleX
        .LIMMINy = .LIMMINy * ScaleY
        .LIMMAXx = .LIMMAXx * ScaleX
        .LIMMAXy = .LIMMAXy * ScaleY
        
        .MinX = .MinX * ScaleX
        .MinY = .MinY * ScaleY
        .MaxX = .MaxX * ScaleX
        .MaxY = .MaxY * ScaleY
    End With
End Sub

Public Sub ScaleEntity(Entity As DXFReaderEntity, ScaleX As Single, ScaleY As Single, ScaleZ As Single)
    
    
    Dim Vertex As DXFReaderVertex
    
    With Entity
        
            If ScaleX <> 1 Then
                .X0 = ScaleX * .X0
                .X1 = ScaleX * .X1
                .X2 = ScaleX * .X2
                If .EntityType <> "IMAGE" Then .X3 = ScaleX * .X3
                If .EntityType = "CIRCLE" Or .EntityType = "ARC" Or .EntityType = "POLYLINE" Or .EntityType = "LWPOLYLINE" Then
                    .g40 = ScaleX * .g40
                End If
                
            End If
            If ScaleY <> 1 Then
                .Y0 = ScaleY * .Y0
                .Y1 = ScaleY * .Y1
                .Y2 = ScaleY * .Y2
                If .EntityType <> "IMAGE" Then .Y3 = ScaleY * .Y3
                If .EntityType = "TEXT" Or .EntityType = "MTEXT" Or .EntityType = "POLYLINE" Or .EntityType = "LWPOLYLINE" Then
                    .g40 = ScaleY * .g40
                End If
                
            End If
            
            If ScaleZ <> 1 Then
                .Z0 = ScaleY * .Z0
                .Z1 = ScaleZ * .Z1
                .z2 = ScaleZ * .z2
                If .EntityType <> "IMAGE" Then .z3 = ScaleZ * .z3
                
            End If
            
            For Each Vertex In .Vertexes
                Vertex.X0 = ScaleX * Vertex.X0
                Vertex.Y0 = ScaleY * Vertex.Y0
                Vertex.Z0 = ScaleY * Vertex.Z0
            Next Vertex
            
       
    End With
End Sub


Private Function TriangleArea(X1, Y1, X2, Y2, X3, Y3) As Double
    TriangleArea = Abs(X1 * (Y2 - Y3) + X2 * (Y3 - Y1) + X3 * (Y1 - Y2)) / 2
End Function
Public Sub CreateNewDrawing(DXFReaderControl As DXFReader)
Attribute CreateNewDrawing.VB_Description = "Creates a new empty drawing."
    With DXFReaderControl
        .NewDrawing
        .CurrentColor = 7
        .FILLMODE = 1
        .BaseX = 0
        .BaseY = 0
        .SetLimits 0, 0, 420, 297
        .ZoomLimits
        .MinX = 0
        .MinY = 0
        .MaxX = 420
        .MaxY = 297
        .FileStatus = 0
        .DateCreated = Now
        .DateLastEdited = Now
    End With
    
End Sub

Public Function FindDistance(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Double
    FindDistance = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
    
End Function

'© Copyright 2000 Kadmos - Software Resolution


Public Function GetLineAngle(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
Attribute GetLineAngle.VB_Description = "Retrieves the angle, in degrees, formed by the x axis and the line from x1,y1 to x2,y2."
    If X1 = X2 Then
        If Y1 = Y2 Then GetLineAngle = -1
        If Y2 > Y1 Then GetLineAngle = 90
        If Y2 < Y1 Then GetLineAngle = 270
    Else
        If Y1 = Y2 Then
            If X1 = X2 Then GetLineAngle = -1
            If X2 > X1 Then GetLineAngle = 0
            If X2 < X1 Then GetLineAngle = 180
        End If
    End If
    If X1 <> X2 And Y2 <> Y1 Then
        If X2 > X1 Then
            If Y2 > Y1 Then
                GetLineAngle = Atn((Y2 - Y1) / (X2 - X1)) * 180 / PI
            Else
                GetLineAngle = -Atn(Abs(Y2 - Y1) / (X2 - X1)) * 180 / PI + 360
            End If
        Else
            If Y2 > Y1 Then
                GetLineAngle = -Atn((Y2 - Y1) / Abs(X2 - X1)) * 180 / PI + 180
            Else
                GetLineAngle = Atn((Y2 - Y1) / (X2 - X1)) * 180 / PI + 180
            End If
        End If
    End If
End Function

Public Sub Add3DFace(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, X2 As Single, Y2 As Single, z2 As Single, X3 As Single, Y3 As Single, z3 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute Add3DFace.VB_Description = "Adds a 3DFace to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "3DFACE"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        .Entity.X1 = X1
        .Entity.Y1 = Y1
        .Entity.Z1 = Z1
        .Entity.X2 = X2
        .Entity.Y2 = Y2
        .Entity.z2 = z2
        .Entity.X3 = X3
        .Entity.Y3 = Y3
        .Entity.z3 = z3
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub




Public Sub AddArc(DXFReaderControl As DXFReader, Xc As Single, Yc As Single, Zc As Single, Radius As Single, StartAngle As Single, EndAngle As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddArc.VB_Description = "Adds an arc to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "ARC"
        .Entity.X0 = Xc
        .Entity.Y0 = Yc
        .Entity.Z0 = Zc
        .Entity.g40 = Radius
        .Entity.g50 = StartAngle
        .Entity.g51 = EndAngle
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub


Public Sub AddEllipse(DXFReaderControl As DXFReader, Xc As Single, Yc As Single, Zc As Single, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddEllipse.VB_Description = "Adds an ellipse to the drawing."
    
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "ELLIPSE"
        .Entity.X0 = Xc
        .Entity.Y0 = Yc
        .Entity.Z0 = Zc
        .Entity.X1 = Xc - X0
        .Entity.Y1 = Yc - Y0
        .Entity.Z1 = Zc - Z0
        If Sqr((Xc - X0) ^ 2 + (Yc - Y0) ^ 2) <> 0 Then
            .Entity.g40 = Sqr((Xc - X1) ^ 2 + (Yc - Y1) ^ 2) / Sqr((Xc - X0) ^ 2 + (Yc - Y0) ^ 2)
        End If
        .Entity.g41 = 0
        .Entity.g42 = 2 * PI
            
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
    
End Sub


Public Sub AddImage(DXFReaderControl As DXFReader, ParentForm As Form, ImageFile As String, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, Optional LineType As String, Optional LayerName As String)
Attribute AddImage.VB_Description = "Adds a bitmap image to the drawing."
    With DXFReaderControl
            .AddEntity
            .Entity.EntityType = "IMAGE"
            .Entity.g70 = 6
            .Entity.X0 = X0
            .Entity.Y0 = Y0
            .Entity.g281 = 50
            .Entity.g282 = 50
            .Entity.g283 = 0
            .Entity.g90 = 2
            .Entity.Z0 = Z0
            .Entity.Z1 = Z1
            .Entity.z2 = 0
            .Entity.z3 = 0
            .Entity.g71 = 1
            Dim ImageHandle As String
            ImageHandle = Hex(Int(Rnd * 100000))
            .Entity.g340 = ImageHandle
            
            .Entity.X3 = ParentForm.ScaleX(LoadPicture(ImageFile).Width, vbHimetric, vbPixels)
            .Entity.Y3 = ParentForm.ScaleY(LoadPicture(ImageFile).Height, vbHimetric, vbPixels)
       
            .Entity.X1 = (X1 - X0) / .Entity.X3
            .Entity.Y1 = 0
            .Entity.X2 = 0
            .Entity.Y2 = (Y1 - Y0) / .Entity.Y3
            
            .AddObject
            .Objects(.Objects.Count).Handle = ImageHandle
            .Objects(.Objects.Count).g1 = ImageFile
            .Objects(.Objects.Count).ObjectType = "IMAGEDEF"
            
            If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
        End With
End Sub

Public Sub AddLine(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddLine.VB_Description = "Add a line to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "LINE"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        .Entity.X1 = X1
        .Entity.Y1 = Y1
        .Entity.Z1 = Z1
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub

Public Sub AddRectangle(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, X1 As Single, Y1 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddRectangle.VB_Description = "Adds a rectangle to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "POLYLINE"
        .Entity.X0 = 0
        .Entity.Y0 = 0
        .Entity.Z0 = 0
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
        
        .Entity.g70 = 1
        
        .AddVertex .Entities.Count
        .Entity.Vertexes(.Entity.Vertexes.Count).X0 = X0
        .Entity.Vertexes(.Entity.Vertexes.Count).Y0 = Y0
        .Entity.Vertexes(.Entity.Vertexes.Count).Z0 = 0
        .AddVertex .Entities.Count
        .Entity.Vertexes(.Entity.Vertexes.Count).X0 = X1
        .Entity.Vertexes(.Entity.Vertexes.Count).Y0 = Y0
        .Entity.Vertexes(.Entity.Vertexes.Count).Z0 = 0
        .AddVertex .Entities.Count
        .Entity.Vertexes(.Entity.Vertexes.Count).X0 = X1
        .Entity.Vertexes(.Entity.Vertexes.Count).Y0 = Y1
        .Entity.Vertexes(.Entity.Vertexes.Count).Z0 = 0
        .AddVertex .Entities.Count
        .Entity.Vertexes(.Entity.Vertexes.Count).X0 = X0
        .Entity.Vertexes(.Entity.Vertexes.Count).Y0 = Y1
        .Entity.Vertexes(.Entity.Vertexes.Count).Z0 = 0
    
    End With
End Sub

Public Sub AddCircle(DXFReaderControl As DXFReader, Xc As Single, Yc As Single, Zc As Single, Radius As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddCircle.VB_Description = "Adds a circle to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "CIRCLE"
        .Entity.X0 = Xc
        .Entity.Y0 = Yc
        .Entity.Z0 = Zc
        .Entity.g40 = Radius
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub
Public Function AddLWPolyLine(DXFReaderControl As DXFReader, Optional Closed = False, Optional Color As Integer, Optional LineType As String, Optional LayerName As String) As Long
Attribute AddLWPolyLine.VB_Description = "Adds a lwpolyline to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "LWPOLYLINE"
        .Entity.X0 = 0
        .Entity.Y0 = 0
        .Entity.Z0 = 0
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        AddLWPolyLine = .Entities.Count
        If Closed Then
            .Entity.g70 = 1
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Function


Public Sub AddPoint(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddPoint.VB_Description = "Adds a point to the drawing."

  With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "POINT"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub

Public Sub InsertBlock(DXFReaderControl As DXFReader, BlockName As String, X0 As Single, Y0 As Single, Z0 As Single, Optional ScaleX As Single = 1, Optional ScaleY As Single = 1, Optional RotationAngle As Single = 0, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute InsertBlock.VB_Description = "Inserts an already defined block in the drawing."
  With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "INSERT"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        .Entity.g41 = ScaleX
        .Entity.g42 = ScaleY
        .Entity.g50 = RotationAngle
        .Entity.BlockName = BlockName
        
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub
Public Function AddPolyLine(DXFReaderControl As DXFReader, Optional Closed = False, Optional Width As Single = 0, Optional Color As Integer, Optional LineType As String, Optional LayerName As String) As Long
Attribute AddPolyLine.VB_Description = "Adds a polyline to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "POLYLINE"
        .Entity.X0 = 0
        .Entity.Y0 = 0
        .Entity.Z0 = 0
        .Entity.g40 = Width
        .Entity.g41 = Width
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        AddPolyLine = .Entities.Count
        If Closed Then
            .Entity.g70 = 1
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Function

Public Sub AddRay(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddRay.VB_Description = "Adds a ray to the drawing."
    
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "RAY"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        
        .Entity.X1 = Cos(GetLineAngle(X0, Y0, X1, Y1) * PI / 180)
        .Entity.Y1 = Sin(GetLineAngle(X0, Y0, X1, Y1) * PI / 180)
        .Entity.Z1 = 0
    
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub
Public Sub AddSolid(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, X2 As Single, Y2 As Single, z2 As Single, X3 As Single, Y3 As Single, z3 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddSolid.VB_Description = "Adds a solid to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "SOLID"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        .Entity.X1 = X1
        .Entity.Y1 = Y1
        .Entity.Z1 = Z1
        .Entity.X2 = X2
        .Entity.Y2 = Y2
        .Entity.z2 = z2
        .Entity.X3 = X3
        .Entity.Y3 = Y3
        .Entity.z3 = z3
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub



Public Sub AddText(DXFReaderControl As DXFReader, Text As String, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, TextHeight As Single, RotationAngle As Single, Optional Color As Integer, Optional TextStyle As String, Optional LineType As String, Optional LayerName As String)
Attribute AddText.VB_Description = "Adds a text to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "TEXT"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        .Entity.X1 = X1
        .Entity.Y1 = Y1
        .Entity.Z1 = Z1
        .Entity.g40 = TextHeight
        .Entity.Text = Text
        .Entity.g50 = RotationAngle
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If TextStyle <> "" Then
            .Entity.TextStyle = TextStyle
        Else
            .Entity.TextStyle = .CurrentTextStyle
        End If
        
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub


Public Sub AddTrace(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, TraceWidth As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddTrace.VB_Description = "Adds a trace to the drawing."
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "TRACE"
        .Entity.X0 = X0 - Cos((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.Y0 = Y0 + Sin((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.Z0 = Z0
        .Entity.X1 = X0 + Cos((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.Y1 = Y0 - Sin((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.Z1 = Z0
        .Entity.X3 = X1 + Cos((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.Y3 = Y1 - Sin((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.z3 = Z1
        .Entity.X2 = X1 - Cos((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.Y2 = Y1 + Sin((90 - GetLineAngle(X0, Y0, X1, Y1)) * PI / 180) * TraceWidth / 2
        .Entity.z2 = Z1
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub

Public Sub AddVertex(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, PolyLineNumber As Long)
Attribute AddVertex.VB_Description = "Adds a vertex to the polyline or the lwpolyline indicated by the PolyLineNumber parameter."
    With DXFReaderControl
        .AddVertex PolyLineNumber
        .Entities(PolyLineNumber).Vertexes(.Entities(PolyLineNumber).Vertexes.Count).X0 = X0
        .Entities(PolyLineNumber).Vertexes(.Entities(PolyLineNumber).Vertexes.Count).Y0 = Y0
        .Entities(PolyLineNumber).Vertexes(.Entities(PolyLineNumber).Vertexes.Count).Z0 = Z0
    End With
End Sub

Public Sub AddXline(DXFReaderControl As DXFReader, X0 As Single, Y0 As Single, Z0 As Single, X1 As Single, Y1 As Single, Z1 As Single, Optional Color As Integer, Optional LineType As String, Optional LayerName As String)
Attribute AddXline.VB_Description = "Adds a construction line to the drawing."
    
    
    With DXFReaderControl
        .AddEntity
        .Entity.EntityType = "XLINE"
        .Entity.X0 = X0
        .Entity.Y0 = Y0
        .Entity.Z0 = Z0
        
        
        .Entity.X1 = Cos(GetLineAngle(X0, Y0, X1, Y1) * PI / 180)
        .Entity.Y1 = Sin(GetLineAngle(X0, Y0, X1, Y1) * PI / 180)
        .Entity.Z1 = 0
        If Color >= 0 Then
            .Entity.Color = Color
        Else
            .Entity.Color = .CurrentColor
        End If
        
        If LineType <> "" Then
            .Entity.LineType = LineType
        Else
            .Entity.LineType = .CurrentLineType
        End If
        If LayerName <> "" Then
            .Entity.LayerName = LayerName
        Else
            .Entity.LayerName = .CurrentLayer
        End If
    End With
End Sub


Public Sub SelectCurrentColor(DXFReaderControl As DXFReader)
Attribute SelectCurrentColor.VB_Description = "Selects interactively the drawing's current color using the default palette dialog."
    DXFReaderControl.PaletteCaption = "Select Current Color"
    DXFReaderControl.CurrentColor = DXFReaderControl.ShowPalette(DXFReaderControl.CurrentColor)
End Sub


Public Sub SetGrid(DXFReaderControl As DXFReader, GridSpacing As Single)
Attribute SetGrid.VB_Description = "Sets and displays the drawing's grid."
   
    With DXFReaderControl
        .ShowGrid = True
        If .ViewPorts(1).ViewPortName = "*ACTIVE" Then
            .ViewPorts(1).GridOnOff = 1
            .ViewPorts(1).GridSpacingX = GridSpacing
            .ViewPorts(1).GridSpacingY = GridSpacing
        End If
    End With
   
End Sub

Public Sub ExplodeBlock(DXFReaderControl As DXFReader, EntityNumber As Long)
Attribute ExplodeBlock.VB_Description = "Explodes a Block."

    Dim k As Long
    Dim p0x, p0y, p1x, p1y, p2x, p2y, p3x, p3y, px, py
    Dim EntityOfBlock As DXFReaderEntity
    
    With DXFReaderControl
        
        For Each EntityOfBlock In .Blocks(.Entities(EntityNumber).BlockName).Entities
            .AddEntity
            .Entity.EntityType = EntityOfBlock.EntityType
            .Entity.g70 = EntityOfBlock.g70
            .Entity.Color = EntityOfBlock.Color
            
            p0x = EntityOfBlock.X0 + .Entities(EntityNumber).X0 - .Blocks(.Entities(EntityNumber).BlockName).X0
            p0y = EntityOfBlock.Y0 + .Entities(EntityNumber).Y0 - .Blocks(.Entities(EntityNumber).BlockName).Y0
            p1x = EntityOfBlock.X1 + .Entities(EntityNumber).X0 - .Blocks(.Entities(EntityNumber).BlockName).X0
    
            p1y = EntityOfBlock.Y1 + .Entities(EntityNumber).Y0 - .Blocks(.Entities(EntityNumber).BlockName).Y0
            p2x = EntityOfBlock.X2 + .Entities(EntityNumber).X0 - .Blocks(.Entities(EntityNumber).BlockName).X0
            p2y = EntityOfBlock.Y2 + .Entities(EntityNumber).Y0 - .Blocks(.Entities(EntityNumber).BlockName).Y0
            p3x = EntityOfBlock.X3 + .Entities(EntityNumber).X0 - .Blocks(.Entities(EntityNumber).BlockName).X0
            p3y = EntityOfBlock.Y3 + .Entities(EntityNumber).Y0 - .Blocks(.Entities(EntityNumber).BlockName).Y0
            
            px = .Entities(EntityNumber).X0
            py = .Entities(EntityNumber).Y0
            
            If .Entities(EntityNumber).g41 <> 0 Then
                p0x = .Entities(EntityNumber).g41 * (p0x - px) + px
                p1x = .Entities(EntityNumber).g41 * (p1x - px) + px
                p2x = .Entities(EntityNumber).g41 * (p2x - px) + px
                p3x = .Entities(EntityNumber).g41 * (p3x - px) + px
            End If
    
            If .Entities(EntityNumber).g42 <> 0 Then
                p0y = .Entities(EntityNumber).g42 * (p0y - py) + py
                p1y = .Entities(EntityNumber).g42 * (p1y - py) + py
                p2y = .Entities(EntityNumber).g42 * (p2y - py) + py
                p3y = .Entities(EntityNumber).g42 * (p3y - py) + py
            End If
        
    
    
    
        Dim b
        
            If .Entities(EntityNumber).g50 <> 0 Then
                
                b = (p0x - px) * Cos(.Entities(EntityNumber).g50 * PI / 180) - (p0y - py) * Sin(.Entities(EntityNumber).g50 * PI / 180) + px
                p0y = (p0x - px) * Sin(.Entities(EntityNumber).g50 * PI / 180) + (p0y - py) * Cos(.Entities(EntityNumber).g50 * PI / 180) + py
                p0x = b
                
                b = (p1x - px) * Cos(.Entities(EntityNumber).g50 * PI / 180) - (p1y - py) * Sin(.Entities(EntityNumber).g50 * PI / 180) + px
                p1y = (p1x - px) * Sin(.Entities(EntityNumber).g50 * PI / 180) + (p1y - py) * Cos(.Entities(EntityNumber).g50 * PI / 180) + py
                p1x = b
                
                b = (p2x - px) * Cos(.Entities(EntityNumber).g50 * PI / 180) - (p2y - py) * Sin(.Entities(EntityNumber).g50 * PI / 180) + px
                p2y = (p2x - px) * Sin(.Entities(EntityNumber).g50 * PI / 180) + (p2y - py) * Cos(.Entities(EntityNumber).g50 * PI / 180) + py
                p2x = b
                
                b = (p3x - px) * Cos(.Entities(EntityNumber).g50 * PI / 180) - (p3y - py) * Sin(.Entities(EntityNumber).g50 * PI / 180) + px
                p3y = (p3x - px) * Sin(.Entities(EntityNumber).g50 * PI / 180) + (p3y - py) * Cos(.Entities(EntityNumber).g50 * PI / 180) + py
                p3x = b
            
            End If
            
            
            .Entity.X0 = p0x
            .Entity.Y0 = p0y
            
            .Entity.X1 = p1x
            .Entity.Y1 = p1y
            
            .Entity.X2 = p2x
            .Entity.Y2 = p2y
            
            .Entity.X3 = p3x
            .Entity.Y3 = p3y
            
            .Entity.g40 = EntityOfBlock.g40
            .Entity.g41 = EntityOfBlock.g41
            .Entity.g42 = EntityOfBlock.g42
            .Entity.g43 = EntityOfBlock.g43
            .Entity.g44 = EntityOfBlock.g44
            .Entity.g45 = EntityOfBlock.g45
            .Entity.g50 = EntityOfBlock.g50
            .Entity.g51 = EntityOfBlock.g51
            .Entity.g71 = EntityOfBlock.g71
            .Entity.g72 = EntityOfBlock.g72
            .Entity.g73 = EntityOfBlock.g73
            .Entity.g74 = EntityOfBlock.g74
            .Entity.g75 = EntityOfBlock.g75
            .Entity.g76 = EntityOfBlock.g76
            .Entity.g90 = EntityOfBlock.g90
            .Entity.g91 = EntityOfBlock.g91
            .Entity.g280 = EntityOfBlock.g280
            .Entity.g281 = EntityOfBlock.g281
            .Entity.g282 = EntityOfBlock.g282
            .Entity.g283 = EntityOfBlock.g283
            .Entity.g340 = EntityOfBlock.g340
            .Entity.Text = EntityOfBlock.Text
            .Entity.TagString = EntityOfBlock.TagString
            .Entity.PromptString = EntityOfBlock.PromptString
            .Entity.AdditionalText = EntityOfBlock.AdditionalText
            .Entity.LayerName = EntityOfBlock.LayerName
            .Entity.Thickness = EntityOfBlock.Thickness
            .Entity.LineType = EntityOfBlock.LineType
            .Entity.Elev = EntityOfBlock.Elev
            .Entity.LineTypeScale = EntityOfBlock.LineTypeScale
            .Entity.Visibility = EntityOfBlock.Visibility
            .Entity.Space = EntityOfBlock.Space
            .Entity.DimensionStyle = EntityOfBlock.DimensionStyle
            .Entity.TextStyle = EntityOfBlock.TextStyle
            .Entity.ExtrusionX = EntityOfBlock.ExtrusionX
            .Entity.ExtrusionY = EntityOfBlock.ExtrusionY
            .Entity.ExtrusionZ = EntityOfBlock.ExtrusionZ
                   
            For k = 1 To EntityOfBlock.Vertexes.Count
                .AddVertex .Entities.Count
                .Entity.Vertexes(k).X0 = EntityOfBlock.Vertexes(k).X0 + .Entities(EntityNumber).X0 - .Blocks(.Entities(EntityNumber).BlockName).X0
                .Entity.Vertexes(k).Y0 = EntityOfBlock.Vertexes(k).Y0 + .Entities(EntityNumber).Y0 - .Blocks(.Entities(EntityNumber).BlockName).Y0
                .Entity.Vertexes(k).Z0 = EntityOfBlock.Vertexes(k).Z0 + .Entities(EntityNumber).Z0 - .Blocks(.Entities(EntityNumber).BlockName).Z0
                
                If .Entities(EntityNumber).g41 <> 0 Then
                    .Entity.Vertexes(k).X0 = .Entities(EntityNumber).g41 * (.Entity.Vertexes(k).X0 - px) + px
                End If
    
                If .Entities(EntityNumber).g42 <> 0 Then
                    .Entity.Vertexes(k).Y0 = .Entities(EntityNumber).g42 * (.Entity.Vertexes(k).Y0 - py) + py
                End If
                
                
                If .Entities(EntityNumber).g50 <> 0 Then
                
                    b = (.Entity.Vertexes(k).X0 - px) * Cos(.Entities(EntityNumber).g50 * PI / 180) - (p0y - py) * Sin(.Entities(EntityNumber).g50 * PI / 180) + px
                    .Entity.Vertexes(k).Y0 = (.Entity.Vertexes(k).X0 - px) * Sin(.Entities(EntityNumber).g50 * PI / 180) + (.Entity.Vertexes(k).Y0 - py) * Cos(.Entities(EntityNumber).g50 * PI / 180) + py
                    .Entity.Vertexes(k).X0 = b
            
                End If
                
                
                .Entity.Vertexes(k).g40 = EntityOfBlock.Vertexes(k).g40
                .Entity.Vertexes(k).g41 = EntityOfBlock.Vertexes(k).g41
                .Entity.Vertexes(k).g42 = EntityOfBlock.Vertexes(k).g42
                .Entity.Vertexes(k).g50 = EntityOfBlock.Vertexes(k).g50
                .Entity.Vertexes(k).g70 = EntityOfBlock.Vertexes(k).g70
                .Entity.Vertexes(k).g71 = EntityOfBlock.Vertexes(k).g71
                .Entity.Vertexes(k).g72 = EntityOfBlock.Vertexes(k).g72
                .Entity.Vertexes(k).g73 = EntityOfBlock.Vertexes(k).g73
                .Entity.Vertexes(k).g74 = EntityOfBlock.Vertexes(k).g74
                .Entity.Vertexes(k).LayerName = EntityOfBlock.Vertexes(k).LayerName
                .Entity.Vertexes(k).Color = EntityOfBlock.Vertexes(k).Color
            Next
                    
        Next EntityOfBlock
                
        .Entities.Remove (EntityNumber)
        
    End With
End Sub
Public Function FindArea(DXFReaderControl As DXFReader, EntityNumber As Long) As Double
    Dim Vertex As DXFReaderVertex
    Dim k As Long
    Dim i As Long
    Dim Area As Double
    
    FindArea = 0
    With DXFReaderControl.Entities(EntityNumber)
        Select Case .EntityType
            Case "POLYLINE", "LWPOLYLINE"
                Area = 0
                For k = 1 To .Vertexes.Count
                    If k = .Vertexes.Count Then
                        i = 1
                    Else
                        i = k + 1
                    End If
                    Area = Area + .Vertexes(k).X0 * .Vertexes(i).Y0 - .Vertexes(i).X0 * .Vertexes(k).Y0
                Next k
                Area = Abs(Area) / 2
            Case "SOLID", "TRACE"
                Area = Abs(TriangleArea(.X0, .Y0, .X1, .Y1, .X3, .Y3) + TriangleArea(.X0, .Y0, .X3, .Y3, .X2, .Y2))
            Case "CIRCLE"
                Area = .g40 * .g40 * PI
        End Select
    End With
    FindArea = Area
    
   
End Function

Public Sub FillPolyLine(DXFReaderControl As DXFReader, EntityNumber As Long, Color As Long)
    Dim Vertex As DXFReaderVertex
    Dim k As Long
    Dim NewBrush As Long
    Dim OldBrush As Long
    
    
    With DXFReaderControl.Entities(EntityNumber)
        Select Case .EntityType
            Case "POLYLINE", "LWPOLYLINE"
                ReDim Points(1 To .Vertexes.Count) As POINTAPI
                For k = 1 To .Vertexes.Count
                    Points(k).x = DXFReaderControl.CoordXtoPixel(.Vertexes(k).X0)
                    Points(k).y = DXFReaderControl.CoordYtoPixel(.Vertexes(k).Y0)
                Next k
               
                
                NewBrush = CreateSolidBrush(Color)
                OldBrush = SelectObject(DXFReaderControl.hDC, NewBrush)
                
                Polygon DXFReaderControl.hDC, Points(1), .Vertexes.Count
                
                OldBrush = SelectObject(DXFReaderControl.hDC, OldBrush)
                DeleteObject NewBrush
                  
                   
    
        End Select
    End With
    
    
   
End Sub
Public Sub FillCircle(DXFReaderControl As DXFReader, EntityNumber As Long, Color As Long)
    Dim NewBrush As Long
    Dim OldBrush As Long
    
    
    With DXFReaderControl.Entities(EntityNumber)
        Select Case .EntityType
            Case "CIRCLE"
                
               
                
                NewBrush = CreateSolidBrush(Color)
                OldBrush = SelectObject(DXFReaderControl.hDC, NewBrush)
                
                Ellipse DXFReaderControl.hDC, DXFReaderControl.CoordXtoPixel(.X0 - .g40), _
                          DXFReaderControl.CoordYtoPixel(.Y0 - .g40), _
                          DXFReaderControl.CoordXtoPixel(.X0 + .g40), _
                          DXFReaderControl.CoordYtoPixel(.Y0 + .g40)
     
                
                
                OldBrush = SelectObject(DXFReaderControl.hDC, OldBrush)
                DeleteObject NewBrush
                  
                   
    
        End Select
    End With
    
    
   
End Sub
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


