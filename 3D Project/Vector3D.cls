VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Double
Public y As Double
Public Z As Double
Public Normal As New Vector3D
Public Lightness As Double
Public Parent As Object3D

Public Function Perspective(theCamera As Camera) As Vector2D
Dim Xw As Double, Yw As Double, Zw As Double, Xe As Double, Ye As Double, Ze As Double
Dim tempVector As New Vector2D
Dim WorldPoint() As Double
Dim Tview() As Double
Dim ViewPoint() As Double
Dim rot() As Double
Dim temparr() As Double

PrepareArray WorldPoint, 1, 4, x, y, Z, 1
With theCamera
    PrepareArray Tview, 4, 4, .Vm11, .Vm12, .Vm13, .Vm14, .Vm21, .Vm22, .Vm23, .Vm24, .Vm31, .Vm32, .Vm33, .Vm34, .Vm41, .Vm42, .Vm43, .Vm44
End With
MultiplyArray WorldPoint, Tview, ViewPoint
Xe = ViewPoint(1, 1)
Ye = ViewPoint(1, 2)
Ze = ViewPoint(1, 3)
If Ze <> 0 Then
    tempVector.x = Int(1000 * (Xe / Ze))
    tempVector.y = Int(1000 * (Ye / Ze))
End If

Set Perspective = tempVector
End Function


