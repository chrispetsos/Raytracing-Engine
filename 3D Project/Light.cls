VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Light"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Rho As Double
Public Theta As Double
Public Phi As Double
Public LightPoint As New Vector3D
Private LightPointD3D As D3DVECTOR
Public LightNum As Integer
Public Parent As World
Public Ambient As Double
Public Intensity As Double
Public CastShadows As Integer

Public Sub Create(theRho As Double, theTheta As Double, thePhi As Double, theAmbient As Double, theIntensity As Double, DoesCastShadows As Integer)
Rho = theRho
Theta = theTheta
Phi = thePhi
Ambient = theAmbient
Intensity = theIntensity
CastShadows = DoesCastShadows
CreateLightMatrix
End Sub
Public Sub AlterRho(Value As Double)
Rho = Rho + Value
If Rho < 0 Then Rho = 0
CreateLightMatrix
End Sub
Public Sub Altertheta(Value As Double)
Theta = Theta + Value
If Theta < 0 Then
    Theta = 360 - Abs(Value)
ElseIf Theta > 360 Then
    Theta = Value
End If
CreateLightMatrix
End Sub
Public Sub AlterPhi(Value As Double)
Phi = Phi + Value
If Phi < 0 Then
    Phi = 360 - Abs(Value)
ElseIf Phi > 360 Then
    Phi = Value
End If
CreateLightMatrix
End Sub
Public Sub CreateLightMatrix()
Dim th As Single, ph As Single, Costh As Single, cosph As Single, sinth As Single, sinph As Single
    
th = DegsToRads(Me.Theta)
ph = DegsToRads(Me.Phi)
Costh = Cos(th)
sinth = Sin(th)
cosph = Cos(ph)
sinph = Sin(ph)
LightPoint.x = Me.Rho * sinph * Costh
LightPoint.y = Me.Rho * sinth * sinph
LightPoint.z = Me.Rho * cosph
VectorToD3D LightPoint, LightPointD3D
End Sub

Public Sub ObjectFaceDistance()
Dim i As Integer
Dim j As Integer

For i = 1 To Me.Parent.NumberOfObjects
    For j = 1 To Me.Parent.GetObject(i).numoffaces
        Me.Parent.GetObject(i).getFace(j).DistanceFromLight = VectorDistance(LightPoint, Me.Parent.GetObject(i).getFace(j).getNode(1))
    Next
Next
End Sub

Public Sub DrawLightOnRaster()
Dim Light2D As New Vector2D
Dim Light2Dradius As Integer
Dim a As Double
Dim YinPicture As Integer
Dim XinPicture As Integer

If CalcLightFromEye <> 0 Then
    Set Light2D = LightPoint.Perspective(Me.Parent.ActiveCamera)
    If (300 / CalcLightFromEye) ^ 4 * 10000 > 200 Then
        Light2Dradius = 200
    Else
        Light2Dradius = (300 / CalcLightFromEye) ^ 4 * 10000
    End If
    Light2D.FixInRasterDevice Me.Parent.RasterDevice
    XinPicture = Light2D.x
    YinPicture = Light2D.y
    If Me.LightNum = Parent.ActiveLight.LightNum Then
        Me.Parent.RasterDevice.Circle (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), Light2Dradius, vbYellow
        Me.Parent.RasterDevice.PSet (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), vbYellow
    Else
        Me.Parent.RasterDevice.Circle (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), Light2Dradius, vbRed
        Me.Parent.RasterDevice.PSet (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), vbRed
    End If
End If
End Sub

Public Function CalcLightFromEye() As Double
CalcLightFromEye = VectorDistance(LightPoint, Parent.ActiveCamera.CameraPoint)
End Function

Public Property Get GetLightPointD3D() As D3DVECTOR
GetLightPointD3D = LightPointD3D
End Property
