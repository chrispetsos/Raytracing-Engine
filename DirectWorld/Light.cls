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
'Public Rho As Double
'Public Theta As Double
'Public Phi As Double
'Public LightPoint As New Vector3D
'Private Position As D3DVECTOR
Private lightD3D As D3DLIGHT8
Public LightNum As Integer
'Public Parent As world
Public Ambient As Double
Public Intensity As Double
Public CastShadows As Integer
Public LightSphere As New CD3DMesh
Public enabled As Boolean

Public Sub Create(thePosition As D3DVECTOR)
Dim retadj As D3DXBuffer

'Rho = theRho
'Theta = theTheta
'Phi = thePhi
'Ambient = theAmbient
'Intensity = theIntensity
lightD3D.Type = D3DLIGHT_POINT
lightD3D.Position = thePosition
lightD3D.Range = 10000
lightD3D.Ambient.r = 0
lightD3D.Ambient.g = 0
lightD3D.Ambient.b = 0
lightD3D.Ambient.a = 1#
lightD3D.diffuse.r = 1#
lightD3D.diffuse.g = 1#
lightD3D.diffuse.b = 1#
lightD3D.diffuse.a = 1#
lightD3D.specular.r = 1#
lightD3D.specular.g = 1#
lightD3D.specular.b = 1#
lightD3D.specular.a = 1#
CastShadows = 0
LightSphere.InitFromD3DXMesh g_d3dx.CreateSphere(g_dev, 20, 20, 20, retadj)
LightSphere.TranslateX thePosition.X
LightSphere.TranslateY thePosition.Y
LightSphere.TranslateZ thePosition.z
enabled = True
'CreateLightMatrix
End Sub
Public Sub AlterX(Value As Single)
lightD3D.Position.X = lightD3D.Position.X + Value
g_dev.SetLight LightNum - 1, lightD3D 'let d3d know about the light
'g_dev.LightEnable LightNum - 1, 1                'turn it on
LightSphere.TranslateX Value
End Sub
Public Sub AlterY(Value As Single)
lightD3D.Position.Y = lightD3D.Position.Y + Value
g_dev.SetLight LightNum - 1, lightD3D 'let d3d know about the light
'g_dev.LightEnable LightNum - 1, 1                'turn it on
LightSphere.TranslateY Value
End Sub
Public Sub AlterZ(Value As Single)
lightD3D.Position.z = lightD3D.Position.z + Value
g_dev.SetLight LightNum - 1, lightD3D 'let d3d know about the light
'g_dev.LightEnable LightNum - 1, 1                'turn it on
LightSphere.TranslateZ Value
End Sub

Public Function GetD3DLight() As D3DLIGHT8
GetD3DLight = lightD3D
End Function

Public Sub SetMaterial(theMaterial As D3DMATERIAL8)
    lightD3D.Ambient = theMaterial.Ambient
    lightD3D.diffuse = theMaterial.diffuse
    lightD3D.specular = theMaterial.specular
End Sub

Public Function GetMaterial() As D3DMATERIAL8
    GetMaterial.Ambient = lightD3D.Ambient
    GetMaterial.diffuse = lightD3D.diffuse
    GetMaterial.specular = lightD3D.specular
End Function

Public Sub draw()
Dim themat As D3DMATERIAL8

'g_dev.SetRenderState D3DRS_AMBIENT, &H202020
If enabled = True Then
'    lightD3D.Ambient.r = 0.1
    If curLight <> LightNum Then
        themat.Ambient.r = 1
    End If
    themat.Ambient.g = 1
    themat.Ambient.b = 1
'    themat.Ambient = lightD3D.Ambient
'    themat.diffuse = lightD3D.diffuse
'    themat.specular = lightD3D.specular
Else
    If curLight <> LightNum Then
        themat.Ambient.r = 0.3
    End If
    themat.Ambient.g = 0.3
    themat.Ambient.b = 0.3
End If
'themat.Ambient.b = 1

LightSphere.bUseMaterials = True
LightSphere.SetMaterial 0, themat
'LightSphere.FlipNormals
LightSphere.Render g_dev
'g_dev.SetRenderState D3DRS_AMBIENT, 0.1
End Sub

Public Sub SetState(theState As Integer)
If theState = 0 Then
    enabled = False
    g_dev.LightEnable LightNum - 1, 0                'turn it off
ElseIf theState = 1 Then
    enabled = True
    g_dev.LightEnable LightNum - 1, 1                'turn it on
End If
End Sub

Public Function GetState() As Integer
If enabled = False Then
    GetState = 0
ElseIf enabled = True Then
    GetState = 1
End If
End Function


'Public Sub CreateLightMatrix()
'Dim th As Single, ph As Single, Costh As Single, cosph As Single, sinth As Single, sinph As Single
'
'th = DegsToRads(Me.Theta)
'ph = DegsToRads(Me.Phi)
'Costh = Cos(th)
'sinth = Sin(th)
'cosph = Cos(ph)
'sinph = Sin(ph)
'LightPoint.X = Me.Rho * sinph * Costh
'LightPoint.Y = Me.Rho * sinth * sinph
'LightPoint.z = Me.Rho * cosph
'VectorToD3D LightPoint, LightPointD3D
'End Sub

'Public Sub ObjectFaceDistance()
'Dim i As Integer
'Dim j As Integer
'
'For i = 1 To Me.Parent.NumberOfObjects
'    For j = 1 To Me.Parent.GetObject(i).numoffaces
'        Me.Parent.GetObject(i).getFace(j).DistanceFromLight = VectorDistance(LightPoint, Me.Parent.GetObject(i).getFace(j).getNode(1))
'    Next
'Next
'End Sub

'Public Sub DrawLightOnRaster()
'Dim Light2D As New Vector2D
'Dim Light2Dradius As Integer
'Dim a As Double
'Dim YinPicture As Integer
'Dim XinPicture As Integer
'
'If CalcLightFromEye <> 0 Then
'    Set Light2D = LightPoint.Perspective(Me.Parent.ActiveCamera)
'    If (300 / CalcLightFromEye) ^ 4 * 10000 > 200 Then
'        Light2Dradius = 200
'    Else
'        Light2Dradius = (300 / CalcLightFromEye) ^ 4 * 10000
'    End If
'    Light2D.FixInRasterDevice Me.Parent.RasterDevice
'    XinPicture = Light2D.X
'    YinPicture = Light2D.Y
'    If Me.LightNum = Parent.ActiveLight.LightNum Then
'        Me.Parent.RasterDevice.Circle (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), Light2Dradius, vbYellow
 ''       Me.Parent.RasterDevice.PSet (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), vbYellow
 '   Else
 '       Me.Parent.RasterDevice.Circle (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), Light2Dradius, vbRed
  ''      Me.Parent.RasterDevice.PSet (FixPointX(Me.Parent.RasterDevice, XinPicture), FixPointY(Me.Parent.RasterDevice, YinPicture)), vbRed
  '  End If'
'End If
'End Sub

'Public Function CalcLightFromEye() As Double
'CalcLightFromEye = VectorDistance(LightPoint, Parent.ActiveCamera.CameraPoint)
'End Function

'Public Property Get GetLightPointD3D() As D3DVECTOR
'GetLightPointD3D = LightPointD3D
'End Property
