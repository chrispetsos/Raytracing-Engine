VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oldCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Rho As Double
Public Theta As Double
Public Phi As Double
Public CameraPoint As New Vector3D
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
Public Parent As World

Public Sub Create(theRho As Double, theTheta As Double, thePhi As Double)
Rho = theRho
Theta = theTheta
Phi = thePhi
'CreateViewMatrix
End Sub
Public Sub AlterRho(Value As Double)
Rho = Rho + Value
If Rho < 0 Then Rho = 0
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub Altertheta(Value As Double)
Theta = Theta + Value
If Theta < 0 Then
    Theta = 360 - Abs(Value)
ElseIf Theta > 360 Then
    Theta = Value
End If
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub AlterPhi(Value As Double)
Phi = Phi + Value
If Phi < 0 Then
    Phi = 360 - Abs(Value)
ElseIf Phi > 360 Then
    Phi = Value
End If
'CreateViewMatrix
ObjectFaceDistance
End Sub
Public Sub CreateViewMatrix()
Dim th As Single, ph As Single, costh As Single, cosph As Single, sinth As Single, sinph As Single
    
th = DegsToRads(Theta)
ph = DegsToRads(Phi)
costh = Cos(th)
sinth = Sin(th)
cosph = Cos(ph)
sinph = Sin(ph)
Vm11 = -sinth
Vm12 = -cosph * costh
Vm13 = -sinph * costh
Vm21 = costh
Vm22 = -cosph * sinth
Vm23 = -sinph * sinth
Vm31 = 0
Vm32 = sinph
Vm33 = -cosph
Vm41 = 0
Vm42 = 0
Vm43 = Rho
CameraPoint.x = Rho * sinph * costh
CameraPoint.y = Rho * sinth * sinph
CameraPoint.Z = Rho * cosph
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

