VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public CameraPoint As New Vector3D
Public FocusPoint As New Vector3D
Public TwistAngle As Double

Public Vm11 As Double
Public Vm12 As Double
Public Vm13 As Double
Public Vm14 As Double
Public Vm21 As Double
Public Vm22 As Double
Public Vm23 As Double
Public Vm24 As Double
Public Vm31 As Double
Public Vm32 As Double
Public Vm33 As Double
Public Vm34 As Double
Public Vm41 As Double
Public Vm42 As Double
Public Vm43 As Double
Public Vm44 As Double
Public CameraNum As Integer
Public Parent As World

Public Sub Create(theCameraPoint As Vector3D, FocusPoint As Vector3D, theTwistAngle As Double)
Set CameraPoint = theCameraPoint
Set N = FocusPoint
TwistAngle = theTwistAngle
CreateViewMatrix
End Sub
Public Sub AlterCameraZ(Value As Double)
CameraPoint.Z = CameraPoint.Z + Value
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub AlterCameraY(Value As Double)
CameraPoint.y = CameraPoint.y + Value
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub AlterCameraX(Value As Double)
CameraPoint.x = CameraPoint.x + Value
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub AlterFocusZ(Value As Double)
FocusPoint.Z = FocusPoint.Z + Value
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub AlterFocusY(Value As Double)
FocusPoint.y = FocusPoint.y + Value
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub AlterFocusX(Value As Double)
FocusPoint.x = FocusPoint.x + Value
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub AlterA(Value As Double)
TwistAngle = TwistAngle + Value
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub CreateViewMatrix()
Dim T1() As Double
Dim T2() As Double
Dim T3() As Double
Dim T4() As Double
Dim T5() As Double
Dim Tview() As Double
Dim temparr1() As Double
Dim temparr2() As Double
Dim temparr3() As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim e As Double
Dim f As Double
Dim g As Double

Dim r As Double
Dim sinth As Double
Dim costh As Double
Dim sinph As Double
Dim cosph As Double
Dim sina As Double
Dim cosa As Double

a = CameraPoint.x
b = CameraPoint.y
c = CameraPoint.Z
e = FocusPoint.x
f = FocusPoint.y
g = FocusPoint.Z

cosa = Cos(DegsToRads(TwistAngle))
sina = Sin(DegsToRads(TwistAngle))
    
If (a = e) And (b = f) Then
    'katakorufh camera
    If c < g Then
        'pros ta katw
        PrepareArray Tview, 4, 4, cosa, -sina, 0, 0, sina, cosa, 0, 0, 0, 0, -1, 0, -a * cosa - b * sina, a * sina - b * cosa, c, 1
    Else
        'pros ta panw
        PrepareArray Tview, 4, 4, -cosa, sina, 0, 0, sina, cosa, 0, 0, 0, 0, 1, 0, a * cosa - b * sina, -a * sina - b * cosa, -c, 1
    End If
Else
    PrepareArray T1, 4, 4, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, -a, -b, -c, 1
    r = Sqr((a - e) ^ 2 + (b - f) ^ 2 + (c - g) ^ 2)
    sinth = (b - f) / Sqr((a - e) ^ 2 + (b - f) ^ 2)
    costh = (a - e) / Sqr((a - e) ^ 2 + (b - f) ^ 2)
    sinph = Sqr((a - e) ^ 2 + (b - f) ^ 2) / r
    cosph = (c - g) / r
    
    PrepareArray T2, 4, 4, sinth, costh, 0, 0, -costh, sinth, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1
    PrepareArray T3, 4, 4, 1, 0, 0, 0, 0, -cosph, -sinph, 0, 0, sinph, -cosph, 0, 0, 0, 0, 1
    PrepareArray T4, 4, 4, -1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1
    
    PrepareArray T5, 4, 4, cosa, -sina, 0, 0, sina, cosa, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1
    
    MultiplyArray T1, T2, temparr1
    MultiplyArray temparr1, T3, temparr2
    MultiplyArray temparr2, T4, temparr3
    MultiplyArray temparr3, T5, Tview
End If
Vm11 = Tview(1, 1)
Vm12 = Tview(1, 2)
Vm13 = Tview(1, 3)
Vm14 = Tview(1, 4)
Vm21 = Tview(2, 1)
Vm22 = Tview(2, 2)
Vm23 = Tview(2, 3)
Vm24 = Tview(2, 4)
Vm31 = Tview(3, 1)
Vm32 = Tview(3, 2)
Vm33 = Tview(3, 3)
Vm34 = Tview(3, 4)
Vm41 = Tview(4, 1)
Vm42 = Tview(4, 2)
Vm43 = Tview(4, 3)
Vm44 = Tview(4, 4)
End Sub

Public Sub ObjectFaceDistance()
Dim i As Integer
Dim j As Integer

For i = 1 To Me.Parent.NumberOfObjects
    For j = 1 To Me.Parent.GetObject(i).NumofFaces
        Me.Parent.GetObject(i).getFace(j).DistanceFromCamera = VectorDistance(CameraPoint, Me.Parent.GetObject(i).getFace(j).getNode(1))
    Next
Next
End Sub
