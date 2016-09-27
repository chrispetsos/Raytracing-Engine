VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "World"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public RasterDevice As PictureBox
Public NumberofCameras As Integer
Public ActiveCamera As New Camera
Private Cameras() As New Camera
Public NumberOfObjects As Integer
Public ActiveObject As New Object3D
Private Objects() As New Object3D
Public NumberofLights As Integer
Public ActiveLight As New Light
Private Lights() As New Light
Private AxisX As New Axis
Private AxisY As New Axis
Private AxisZ As New Axis
Private SortedObjectList() As Integer
Public VertexNormals As Boolean
Public FaceNormals As Boolean
Public Mesh As Boolean
Public Flat As Boolean
Public Gouraud As Boolean
Public Phong As Boolean

Public Sub Create(TheRasterDevice As PictureBox)
Set RasterDevice = TheRasterDevice
NumberofCameras = 0
Set ActiveCamera = Nothing
NumberOfObjects = 0
Set ActiveObject = Nothing
NumberofLights = 0
Set ActiveLight = Nothing
AxisX.EndOfAxis.x = 600
AxisX.EndOfAxis.y = 0
AxisX.EndOfAxis.Z = 0
AxisX.Name = "X"

AxisY.EndOfAxis.x = 0
AxisY.EndOfAxis.y = 600
AxisY.EndOfAxis.Z = 0
AxisY.Name = "Y"

AxisZ.EndOfAxis.x = 0
AxisZ.EndOfAxis.y = 0
AxisZ.EndOfAxis.Z = 600
AxisZ.Name = "Z"

Set AxisX.Parent = Me
Set AxisY.Parent = Me
Set AxisZ.Parent = Me
End Sub

Public Sub AddCamera(theCamera As Camera)
NumberofCameras = NumberofCameras + 1
ReDim Preserve Cameras(1 To NumberofCameras)
Set Cameras(NumberofCameras) = theCamera
Set ActiveCamera = theCamera
Set theCamera.Parent = Me
theCamera.CameraNum = NumberofCameras
End Sub
Public Sub AddObject(theObject As Object3D)
NumberOfObjects = NumberOfObjects + 1
ReDim Preserve Objects(1 To NumberOfObjects)
ReDim Preserve SortedObjectList(1 To NumberOfObjects)
Set Objects(NumberOfObjects) = theObject
Set ActiveObject = theObject
Set theObject.Parent = Me
theObject.objectNum = NumberOfObjects
theObject.CalcVertexInFaces         'Mono mia fora upologizontai edw
ActiveCamera.ObjectFaceDistance
End Sub
Public Sub DeleteObject(objectNum As Integer)
For i = objectNum To NumberOfObjects - 1
    Set Objects(i) = Objects(i + 1)
Next
NumberOfObjects = NumberOfObjects - 1
If NumberOfObjects > 0 Then
    ReDim Preserve Objects(1 To NumberOfObjects)
    ReDim Preserve SortedObjectList(1 To NumberOfObjects)
    If NumberOfObjects = 1 Then
        Set ActiveObject = Objects(1)
    Else
        Set ActiveObject = Objects(objectNum)
    End If
End If
End Sub
Public Sub DeleteCamera(CameraNum As Integer)
For i = CameraNum To NumberofCameras - 1
    Set Cameras(i) = Cameras(i + 1)
    Cameras(i).CameraNum = Cameras(i).CameraNum - 1
Next
NumberofCameras = NumberofCameras - 1
ReDim Preserve Cameras(1 To NumberofCameras)
If NumberofCameras = 1 Then
    Set ActiveCamera = Cameras(1)
ElseIf NumberofCameras = CameraNum - 1 Then
    Set ActiveCamera = Cameras(CameraNum - 1)
Else
    Set ActiveCamera = Cameras(CameraNum)
End If
End Sub
Public Sub DeleteLight(LightNum As Integer)
For i = LightNum To NumberofLights - 1
    Set Lights(i) = Lights(i + 1)
    Lights(i).LightNum = Lights(i).LightNum - 1
Next
NumberofLights = NumberofLights - 1
ReDim Preserve Lights(1 To NumberofLights)
If NumberofLights = 1 Then
    Set ActiveLight = Lights(1)
ElseIf NumberofLights = LightNum - 1 Then
    Set ActiveLight = Lights(LightNum - 1)
Else
    Set ActiveLight = Lights(LightNum)
End If
End Sub
Public Sub AddLight(theLight As Light)
NumberofLights = NumberofLights + 1
ReDim Preserve Lights(1 To NumberofLights)
Set Lights(NumberofLights) = theLight
Set ActiveLight = theLight
Set theLight.Parent = Me
theLight.LightNum = NumberofLights
End Sub
Public Sub Raster(theVertexNormals As Boolean, theFaceNormals As Boolean, theMesh As Boolean, theFlat As Boolean, theGouraud As Boolean, thePhong As Boolean, theAxis As Boolean, theLight As Boolean)
Dim minObjDist As Double

Screen.MousePointer = vbHourglass
VertexNormals = theVertexNormals
FaceNormals = theFaceNormals
Mesh = theMesh
Flat = theFlat
Gouraud = theGouraud
Phong = thePhong

RasterDevice.Cls
ActiveCamera.CreateViewMatrix
If theAxis Then
    AxisX.draw
    AxisY.draw
    AxisZ.draw
End If
If NumberOfObjects > 0 Then
    SortObjectListFromActiveCamera
    minObjDist = Objects(SortedObjectList(1)).DistanceFromCamera
End If
If minObjDist > ActiveLight.CalcLightFromEye Then
    For i = NumberOfObjects To 1 Step -1
        Objects(SortedObjectList(i)).draw
    Next
    If theLight Then
        For i = 1 To NumberofLights
            'Me.ActiveLight.DrawLightOnRaster
            Lights(i).DrawLightOnRaster
        Next
    End If
Else
    If theLight Then
        For i = 1 To NumberofLights
            'Me.ActiveLight.DrawLightOnRaster
            Lights(i).DrawLightOnRaster
        Next
    End If
    For i = NumberOfObjects To 1 Step -1
        Objects(SortedObjectList(i)).draw
    Next
End If
Screen.MousePointer = vbArrow
End Sub

Private Sub SortObjectListFromActiveCamera()
Dim ObjectDepth() As Double
ReDim ObjectDepth(1 To NumberOfObjects)
Dim Min As Double, temp1 As Double, minrow As Integer
Dim temp2 As Double, temp3 As Double, temp4 As Double, k As Integer
For i = 1 To NumberOfObjects
    SortedObjectList(i) = i
    ObjectDepth(i) = Objects(i).DistanceFromCamera
Next
For i = 1 To NumberOfObjects
    Min = ObjectDepth(i)
    k = 0
    For j = i To NumberOfObjects
        If Min > ObjectDepth(j) Then
            k = k + 1
            Min = ObjectDepth(j)
            minrow = j
        ElseIf k = 0 Then
            minrow = i
        End If
    Next j
    temp3 = ObjectDepth(i)
    ObjectDepth(i) = ObjectDepth(minrow)
    ObjectDepth(minrow) = temp3
    temp3 = SortedObjectList(i)
    SortedObjectList(i) = SortedObjectList(minrow)
    SortedObjectList(minrow) = temp3
Next i
End Sub

Public Sub ClickMe(ByVal x As Integer, ByVal y As Integer)
Dim minX As Long
Dim maxX As Long
Dim minY As Long
Dim maxY As Long

Raster VertexNormals, FaceNormals, Mesh, Flat, Gouraud, Phong, True, True
x = x - RasterDevice.ScaleWidth / 2
y = RasterDevice.ScaleWidth / 2 - y
minX = ActiveObject.minX
maxX = ActiveObject.maxX
minY = ActiveObject.minY
maxY = ActiveObject.maxY
For i = 1 To NumberOfObjects
    minX = Objects(SortedObjectList(i)).minX
    maxX = Objects(SortedObjectList(i)).maxX
    minY = Objects(SortedObjectList(i)).minY
    maxY = Objects(SortedObjectList(i)).maxY
    If PointInRectangle(x, y, minX, maxX, minY, maxY) Then
        Set ActiveObject = Objects(SortedObjectList(i))
        Exit For
    End If
Next
MainForm.Label16(2).Caption = ActiveObject.Kd
MainForm.Label16(3).Caption = ActiveObject.Ks
MainForm.Label16(4).Caption = ActiveObject.n
MainForm.Label16(5).Caption = ActiveObject.Hue
MainForm.Label16(6).Caption = ActiveObject.Saturation
MainForm.Label16(7).Caption = ActiveObject.Lightness
RasterDevice.Line (FixPointX(RasterDevice, minX), FixPointY(RasterDevice, maxY))-(FixPointX(RasterDevice, maxX), FixPointY(RasterDevice, minY)), vbRed, B
End Sub


Public Property Get GetObject(objectNum As Integer) As Object3D
    Set GetObject = Objects(objectNum)
End Property
Public Property Get GetCamera(CameraNum As Integer) As Camera
    Set GetCamera = Cameras(CameraNum)
End Property
Public Property Get GetLight(LightNum As Integer) As Light
    Set GetLight = Lights(LightNum)
End Property

