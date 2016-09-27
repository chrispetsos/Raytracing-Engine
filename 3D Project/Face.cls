VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Face"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public NumofNodes As Integer
Private Nodes() As Integer
Private ScreenFace() As New Vector2D
Public Parent As Object3D
Public Color As New FaceColor
Public DistanceFromCamera As Double
Public DistanceFromLight As Double
Private EyeCosAngle As Double
Private lightcosangle As Double
Public NormalVector As New Vector3D
Private CameraVector As New Vector3D
Private LightVector As New Vector3D
Public center As New Vector3D
Private Vector1 As New Vector3D
Private Vector2 As New Vector3D
Private myPoints() As Mpoint


Public Sub Create(NumberofNodes As Integer, ParamArray theNodes())
NumofNodes = NumberofNodes
ReDim Nodes(1 To NumberofNodes)
ReDim ScreenFace(1 To NumberofNodes)

For i = 0 To NumberofNodes - 1
    If UBound(theNodes, 1) <> 0 Then
        Nodes(i + 1) = theNodes(i)
    ElseIf UBound(theNodes(0), 1) <> 0 Then
        Nodes(i + 1) = theNodes(0)(i)
    ElseIf UBound(theNodes(0)(0), 1) <> 0 Then
        Nodes(i + 1) = theNodes(0)(0)(i + 1)
    End If
Next
End Sub

Private Sub CalcNormal()
Set NormalVector = CrossProduct(Vector1, Vector2)
normalize NormalVector
End Sub
Private Sub CalcCameraVector()
Set CameraVector = VectorMinus(Me.Parent.Parent.ActiveCamera.CameraPoint, center)
normalize CameraVector
End Sub
Private Sub CalcLightVector()
Set LightVector = VectorMinus(Me.Parent.Parent.ActiveLight.LightPoint, center)
normalize LightVector
End Sub
Private Sub CalcFaceVectors()
CalcCenter
Set Vector1 = VectorMinus(Me.Parent.getNode(Nodes(3)), center)
Set Vector2 = VectorMinus(Me.Parent.getNode(Nodes(2)), center)
normalize Vector1
normalize Vector2

End Sub
Private Sub CalcEyeAngle()
CalcCameraVector
EyeCosAngle = VectorCosAngle(NormalVector, CameraVector)
End Sub
Private Sub CalcLightAngle()
CalcLightVector
lightcosangle = VectorCosAngle(NormalVector, LightVector)
End Sub
Public Sub CalcBasicVectors()
CalcFaceVectors
CalcNormal
End Sub
Public Function Hidden() As Boolean
If EyeCosAngle > 0 Then
    Hidden = False
Else
    Hidden = True
End If
End Function

Public Sub draw()
Dim Lightness As Double
Dim tempR As Single
Dim tempG As Single
Dim tempB As Single
Dim tempH As Double
Dim tempL As Double
Dim tempS As Double
Dim Intensity(1 To 3) As Double
Dim NodesArray(1 To 3) As New Vector3D
Dim Normals(1 To 3) As New Vector3D
Dim i As Integer

CalcEyeAngle
If Hidden = False Then      'an einai orath
    For i = 1 To NumofNodes
        Set ScreenFace(i) = Me.Parent.getNode(Nodes(i)).Perspective(Me.Parent.Parent.ActiveCamera)
    Next
    ReDim myPoints(1 To Me.NumofNodes)
    For i = 1 To Me.NumofNodes
        myPoints(i).NodeNo = Nodes(i)
        myPoints(i).x = ScreenFace(i).x
        myPoints(i).y = ScreenFace(i).y
    Next
    If Parent.Parent.Flat Then      'Flat Shading
        Index = Parent.Parent.ActiveLight.LightNum
        Lightness = 0
        For i = 1 To Parent.Parent.NumberofLights
            Set Parent.Parent.ActiveLight = Parent.Parent.GetLight(i)
            CalcLightAngle
            If lightcosangle > 0 Then
                Lightness = Lightness + Parent.Parent.ActiveLight.Ambient + Parent.Parent.ActiveLight.Intensity * Parent.Kd * lightcosangle
            Else
                Lightness = Lightness + Parent.Parent.ActiveLight.Ambient
            End If
        Next
'        RGB2HLS Parent.Color.r / 255, Parent.Color.g / 255, Parent.Color.b / 255, tempH, tempL, tempS
        Lightness = Lightness + Parent.Lightness
        If Lightness < 0 Then Lightness = 0
        If Lightness > 1 Then Lightness = 1
        hls2rgb Parent.Hue, Lightness, Parent.Saturation, tempR, tempG, tempB
        Color.r = tempR * 255
        Color.g = tempG * 255
        Color.b = tempB * 255
        tempR = Color.r
        tempG = Color.g
        tempB = Color.b
        doFlat Parent.Parent.RasterDevice, myPoints, RGB(tempR, tempG, tempB)
        Parent.Parent.ActiveLight.LightNum = Index
    ElseIf Parent.Parent.Gouraud Then       'Gouraud Shading
        Intensity(1) = Parent.getNode(Nodes(1)).Lightness
        Intensity(2) = Parent.getNode(Nodes(2)).Lightness
        Intensity(3) = Parent.getNode(Nodes(3)).Lightness
        doGouraud Parent.Parent.RasterDevice, myPoints, Intensity, Me
    ElseIf Parent.Parent.Phong Then       'Phong Shading
        Set NodesArray(1) = Parent.getNode(Nodes(1))
        Set NodesArray(2) = Parent.getNode(Nodes(2))
        Set NodesArray(3) = Parent.getNode(Nodes(3))
        Set NodesArray(1).Parent = Parent
        Set NodesArray(2).Parent = Parent
        Set NodesArray(3).Parent = Parent
        Set Normals(1) = Parent.getNode(Nodes(1)).Normal
        Set Normals(2) = Parent.getNode(Nodes(2)).Normal
        Set Normals(3) = Parent.getNode(Nodes(3)).Normal
        doPhong Parent.Parent, myPoints, NodesArray, Normals, Color
    End If
    If Parent.Parent.Mesh Then      'An thelei Mesh
        drawMesh
    End If
End If
End Sub
Public Sub drawMesh()
Parent.Parent.RasterDevice.Line (FixPointX(Parent.Parent.RasterDevice, myPoints(1).x), FixPointY(Parent.Parent.RasterDevice, myPoints(1).y))-(FixPointX(Parent.Parent.RasterDevice, myPoints(2).x), FixPointY(Parent.Parent.RasterDevice, myPoints(2).y)), vbMagenta
Parent.Parent.RasterDevice.Line (FixPointX(Parent.Parent.RasterDevice, myPoints(3).x), FixPointY(Parent.Parent.RasterDevice, myPoints(3).y))-(FixPointX(Parent.Parent.RasterDevice, myPoints(2).x), FixPointY(Parent.Parent.RasterDevice, myPoints(2).y)), vbMagenta
Parent.Parent.RasterDevice.Line (FixPointX(Parent.Parent.RasterDevice, myPoints(1).x), FixPointY(Parent.Parent.RasterDevice, myPoints(1).y))-(FixPointX(Parent.Parent.RasterDevice, myPoints(3).x), FixPointY(Parent.Parent.RasterDevice, myPoints(3).y)), vbMagenta
End Sub

Public Property Get getNode(NodeNo) As Vector3D
    Set getNode = Me.Parent.getNode(Nodes(NodeNo))
End Property

Public Property Get getNodeNo(NodeNo) As Integer
    getNodeNo = Nodes(NodeNo)
End Property

Public Sub CalcCenter()
Dim totalX As Double
Dim totalY As Double
Dim totalZ As Double

totalX = 0
totalY = 0
totalZ = 0
For i = 1 To NumofNodes
    totalX = totalX + Parent.getNode(Nodes(i)).x
    totalY = totalY + Parent.getNode(Nodes(i)).y
    totalZ = totalZ + Parent.getNode(Nodes(i)).Z
Next
center.x = totalX / NumofNodes
center.y = totalY / NumofNodes
center.Z = totalZ / NumofNodes
End Sub

Public Sub drawNormal()
Dim theNormal As New Vector3D
Dim NormalStart As New Vector2D
Dim NormalEnd As New Vector2D

If Hidden = False Then
    If Parent.Parent.FaceNormals Then
        theNormal.x = NormalVector.x * 50
        theNormal.y = NormalVector.y * 50
        theNormal.Z = NormalVector.Z * 50
        
        theNormal.x = theNormal.x + center.x
        theNormal.y = theNormal.y + center.y
        theNormal.Z = theNormal.Z + center.Z
        
        Set NormalStart = center.Perspective(Parent.Parent.ActiveCamera)
        Set NormalEnd = theNormal.Perspective(Parent.Parent.ActiveCamera)
        NormalStart.x = FixPointX(Parent.Parent.RasterDevice, NormalStart.x)
        NormalStart.y = FixPointY(Parent.Parent.RasterDevice, NormalStart.y)
        NormalEnd.x = FixPointX(Parent.Parent.RasterDevice, NormalEnd.x)
        NormalEnd.y = FixPointY(Parent.Parent.RasterDevice, NormalEnd.y)
        
        Parent.Parent.RasterDevice.Line (NormalStart.x, NormalStart.y)-(NormalEnd.x, NormalEnd.y), vbRed
    End If
    If Parent.Parent.VertexNormals Then
        drawVertexNormals
    End If
End If
End Sub

Public Sub drawVertexNormals()
Dim theNormal As New Vector3D
Dim NormalStart As New Vector2D
Dim NormalEnd As New Vector2D

For i = 1 To NumofNodes
    theNormal.x = Parent.getNode(Nodes(i)).Normal.x * 50
    theNormal.y = Parent.getNode(Nodes(i)).Normal.y * 50
    theNormal.Z = Parent.getNode(Nodes(i)).Normal.Z * 50
    
    theNormal.x = theNormal.x + Parent.getNode(Nodes(i)).x
    theNormal.y = theNormal.y + Parent.getNode(Nodes(i)).y
    theNormal.Z = theNormal.Z + Parent.getNode(Nodes(i)).Z
    
    Set NormalStart = Parent.getNode(Nodes(i)).Perspective(Parent.Parent.ActiveCamera)
    Set NormalEnd = theNormal.Perspective(Parent.Parent.ActiveCamera)
    NormalStart.x = FixPointX(Parent.Parent.RasterDevice, NormalStart.x)
    NormalStart.y = FixPointY(Parent.Parent.RasterDevice, NormalStart.y)
    NormalEnd.x = FixPointX(Parent.Parent.RasterDevice, NormalEnd.x)
    NormalEnd.y = FixPointY(Parent.Parent.RasterDevice, NormalEnd.y)
    
    Parent.Parent.RasterDevice.Line (NormalStart.x, NormalStart.y)-(NormalEnd.x, NormalEnd.y), vbYellow
Next
End Sub

Public Property Get getScreenFaceNode(NodeNo As Object) As Vector2D
    Set getScreenFaceNode = ScreenFace(NodeNo)
End Property

Public Property Get minX() As Long
minX = ScreenFace(1).x
For i = 2 To NumofNodes
    If ScreenFace(i).x < minX Then
        minX = ScreenFace(i).x
    End If
Next
End Property
Public Property Get maxX() As Long
maxX = ScreenFace(1).x
For i = 2 To NumofNodes
    If ScreenFace(i).x > maxX Then
        maxX = ScreenFace(i).x
    End If
Next
End Property
Public Property Get minY() As Long
minY = ScreenFace(1).y
For i = 2 To NumofNodes
    If ScreenFace(i).y < minY Then
        minY = ScreenFace(i).y
    End If
Next
End Property
Public Property Get maxY() As Long
maxY = ScreenFace(1).y
For i = 2 To NumofNodes
    If ScreenFace(i).y > maxY Then
        maxY = ScreenFace(i).y
    End If
Next
End Property

