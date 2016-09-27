VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Object3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public objectNum As Integer
Public NumofNodes As Integer
Public NumofFaces As Integer
Private Faces() As New Face
Private PrimalNodes() As New Vector3D
Private Nodes() As New Vector3D
Private VertexInFaces() As Variant
Public Translation As New Vector3D
Public Rotation As New Vector3D
Public Scaling As New Vector3D
Public WorldRef As New Vector3D
Public Parent As New World
Public Color As New FaceColor
Public DistanceFromCamera As Double
Public Hue As Integer
Public Lightness As Double
Public Saturation As Double
Public Kd As Double
Public Ks As Double
Public n As Double
Private SortedFaceList() As Integer

Public Sub Create(NumberofNodes As Integer, NumberOfFaces As Integer, theHue As Integer, theLightness As Double, theSaturation As Double, theKd As Double, theKs As Double, theNi As Double)
NumofNodes = NumberofNodes
NumofFaces = NumberOfFaces
ReDim PrimalNodes(1 To NumberofNodes)
ReDim Nodes(1 To NumberofNodes)
ReDim VertexInFaces(1 To NumberofNodes)
ReDim Faces(1 To NumberOfFaces)
ReDim SortedFaceList(1 To NumberOfFaces)

Translation.x = General.ZeroVector.x
Translation.y = General.ZeroVector.y
Translation.Z = General.ZeroVector.Z
Rotation.x = General.ZeroVector.x
Rotation.y = General.ZeroVector.y
Rotation.Z = General.ZeroVector.Z
Scaling.x = General.ZeroVector.x
Scaling.y = General.ZeroVector.y
Scaling.Z = General.ZeroVector.Z
WorldRef.x = General.ZeroVector.x
WorldRef.y = General.ZeroVector.y
WorldRef.Z = General.ZeroVector.Z
'Set Color = theColor
Hue = theHue
Lightness = theLightness
Saturation = theSaturation
Kd = theKd
Ks = theKs
n = theNi
End Sub
Public Sub AddNode(ByVal NodeNumber As Integer, theX As Double, theY As Double, theZ As Double)
Nodes(NodeNumber).x = theX
Nodes(NodeNumber).y = theY
Nodes(NodeNumber).Z = theZ
PrimalNodes(NodeNumber).x = theX
PrimalNodes(NodeNumber).y = theY
PrimalNodes(NodeNumber).Z = theZ
Set PrimalNodes(NodeNumber).Parent = Me
End Sub
Public Sub AddFace(FaceNumber As Integer, NumberofNodes As Integer, ParamArray theNodes())
Faces(FaceNumber).Create NumberofNodes, theNodes
Set Faces(FaceNumber).Parent = Me
Set Faces(FaceNumber).Color = Me.Color
End Sub
Public Sub TranslateIt()
For i = 1 To NumofNodes
    Nodes(i).x = Nodes(i).x + Translation.x
    Nodes(i).y = Nodes(i).y + Translation.y
    Nodes(i).Z = Nodes(i).Z + Translation.Z
Next
MoveRefPoint Translation
End Sub
Public Sub TranslateX(ByVal Value As Double)
Translation.x = Translation.x + Value
End Sub
Public Sub TranslateY(ByVal Value As Double)
Translation.y = Translation.y + Value
End Sub
Public Sub TranslateZ(ByVal Value As Double)
Translation.Z = Translation.Z + Value
End Sub
Public Sub RotateIt()
Dim i As Integer
Dim tempX As Double
Dim tempY As Double
Dim tempZ As Double
Dim rotangle As Double

rotangle = DegsToRads(Rotation.x)

For i = 1 To NumofNodes
    tempX = Nodes(i).x
    tempY = Nodes(i).y
    tempZ = Nodes(i).Z
    Nodes(i).y = tempY * Cos(rotangle) - tempZ * Sin(rotangle)
    Nodes(i).Z = tempY * Sin(rotangle) + tempZ * Cos(rotangle)
Next

rotangle = DegsToRads(Rotation.y)

For i = 1 To NumofNodes
    tempX = Nodes(i).x
    tempY = Nodes(i).y
    tempZ = Nodes(i).Z
    Nodes(i).x = tempX * Cos(rotangle) + tempZ * Sin(rotangle)
    Nodes(i).Z = -tempX * Sin(rotangle) + tempZ * Cos(rotangle)
Next
rotangle = DegsToRads(Rotation.Z)

For i = 1 To NumofNodes
    tempX = Nodes(i).x
    tempY = Nodes(i).y
    tempZ = Nodes(i).Z
    Nodes(i).x = tempX * Cos(rotangle) - tempY * Sin(rotangle)
    Nodes(i).y = tempX * Sin(rotangle) + tempY * Cos(rotangle)
Next
End Sub
Public Sub rotateX(ByVal val As Integer)
Rotation.x = Rotation.x + val
End Sub

Public Sub rotatey(ByVal val As Integer)
Rotation.y = Rotation.y + val
End Sub

Public Sub rotateZ(ByVal val As Integer)
Rotation.Z = Rotation.Z + val
End Sub
Public Sub ScaleIt()
For i = 1 To NumofNodes
'    Nodes(i).x = PrimalNodes(i).x + PrimalNodes(i).x * Scaling.x
'    Nodes(i).y = PrimalNodes(i).y + PrimalNodes(i).y * Scaling.y
'    Nodes(i).Z = PrimalNodes(i).Z + PrimalNodes(i).Z * Scaling.Z
    Nodes(i).x = Nodes(i).x + Nodes(i).x * Scaling.x
    Nodes(i).y = Nodes(i).y + Nodes(i).y * Scaling.y
    Nodes(i).Z = Nodes(i).Z + Nodes(i).Z * Scaling.Z
Next
End Sub
Public Sub ScaleX(ByVal val As Double)
Scaling.x = Scaling.x + val
End Sub

Public Sub ScaleY(ByVal val As Double)
Scaling.y = Scaling.y + val
End Sub

Public Sub ScaleZ(ByVal val As Double)
Scaling.Z = Scaling.Z + val
End Sub
Public Sub MoveRefPoint(MoveVector As Vector3D)
WorldRef.x = MoveVector.x
WorldRef.y = MoveVector.y
WorldRef.Z = MoveVector.Z
End Sub
Public Sub AssignPrimalNodes()
For i = 1 To NumofNodes
    Nodes(i).x = PrimalNodes(i).x
    Nodes(i).y = PrimalNodes(i).y
    Nodes(i).Z = PrimalNodes(i).Z
Next
End Sub

Public Sub draw()
AssignPrimalNodes
If Not IsZeroVector(Scaling) Then ScaleIt     'Object Translations
If Not IsZeroVector(Rotation) Then RotateIt
If Not IsZeroVector(Translation) Then TranslateIt

For i = NumofFaces To 1 Step -1
    Faces(i).CalcBasicVectors   'Gia ton Hidden algorithm
Next

'An xreiastoun ta Vertex Normals upologise ta
If Parent.VertexNormals = True Or Parent.FaceNormals = True Or Parent.Gouraud Or Parent.Phong Then
    CalcVertexNormals
End If

CalcDistanceFromCamera      'gia to an tha graftei prwta to fws h oxi

If Parent.Flat Then             'Flat Shading
    CalcFacesDistances          'Painter's algorithm
    SortFaceListFromActiveCamera
    For i = NumofFaces To 1 Step -1
        Faces(SortedFaceList(i)).draw
    Next
ElseIf Parent.Gouraud Then      'Gouraud shadnig
    CalcFacesDistances          'Painter's algorithm
    SortFaceListFromActiveCamera
    CalcVertexLightness         'Tha xreiastei gia to gouraud shading
    For i = NumofFaces To 1 Step -1
        Faces(SortedFaceList(i)).draw
    Next
ElseIf Parent.Phong Then      'Phong shadnig
    CalcFacesDistances          'Painter's algorithm
    SortFaceListFromActiveCamera
'    CalcVertexLightness         'Tha xreiastei gia to gouraud shading
    For i = NumofFaces To 1 Step -1
        Faces(SortedFaceList(i)).draw
    Next
Else                            'No Shading
    For i = NumofFaces To 1 Step -1
        Faces(i).draw
    Next
End If
                                'An tha emfanistoun normals
If Parent.VertexNormals = True Or Parent.FaceNormals = True Then
    For i = NumofFaces To 1 Step -1
        Faces(i).drawNormal
    Next
End If
End Sub
Private Sub SortFaceListFromActiveCamera()
Dim FaceDepth() As Double
ReDim FaceDepth(1 To NumofFaces)
Dim Min As Double, temp1 As Double, minrow As Integer
Dim temp2 As Double, temp3 As Double, temp4 As Double, k As Integer
For i = 1 To NumofFaces
    SortedFaceList(i) = i
    FaceDepth(i) = Faces(i).DistanceFromCamera
Next
For i = 1 To NumofFaces
    Min = FaceDepth(i)
    k = 0
    For j = i To NumofFaces
        If Min > FaceDepth(j) Then
            k = k + 1
            Min = FaceDepth(j)
            minrow = j
        ElseIf k = 0 Then
            minrow = i
        End If
    Next j
    temp3 = FaceDepth(i)
    FaceDepth(i) = FaceDepth(minrow)
    FaceDepth(minrow) = temp3
    temp3 = SortedFaceList(i)
    SortedFaceList(i) = SortedFaceList(minrow)
    SortedFaceList(minrow) = temp3
Next i
End Sub

Public Property Get getNode(NodeNo) As Vector3D
    Set getNode = Nodes(NodeNo)
End Property
Public Property Get getFace(FaceNo) As Face
    Set getFace = Faces(FaceNo)
End Property

Public Sub CalcFacesDistances()
For i = 1 To NumofFaces
    Faces(i).DistanceFromCamera = VectorDistance(Me.Parent.ActiveCamera.CameraPoint, Faces(i).center)
Next
End Sub

Public Sub CalcDistanceFromCamera()
DistanceFromCamera = VectorDistance(Me.Parent.ActiveCamera.CameraPoint, WorldRef)
End Sub

Public Sub CalcVertexInFaces()
Dim temparr() As Integer
Dim col As Integer

col = 0
For i = 1 To NumofNodes
    For j = 1 To NumofFaces
        For k = 1 To Faces(j).NumofNodes
            If Faces(j).getNodeNo(k) = i Then
                col = col + 1
                ReDim Preserve temparr(1 To col)
                temparr(col) = j
                Exit For
            End If
        Next
    Next
    VertexInFaces(i) = temparr()
    ReDim temparr(1 To 1)
    col = 0
Next
End Sub
Public Sub CalcVertexNormals()
Dim totalX As Double
Dim totalY As Double
Dim totalZ As Double
Dim ParticipatesIn As Integer

For i = 1 To NumofNodes
    ParticipatesIn = UBound(VertexInFaces(i))
    totalX = 0
    totalY = 0
    totalZ = 0
    For j = 1 To ParticipatesIn
        totalX = totalX + Faces(VertexInFaces(i)(j)).NormalVector.x
        totalY = totalY + Faces(VertexInFaces(i)(j)).NormalVector.y
        totalZ = totalZ + Faces(VertexInFaces(i)(j)).NormalVector.Z
    Next
    Nodes(i).Normal.x = totalX / ParticipatesIn
    Nodes(i).Normal.y = totalY / ParticipatesIn
    Nodes(i).Normal.Z = totalZ / ParticipatesIn
Next
End Sub

Public Sub CalcVertexLightness()
Dim lightcosangle As Double
Dim mirrorcosangle As Double
Dim LightP As New Vector3D
Dim MirrorP As New Vector3D
Dim nodenormal As New Vector3D
Dim tempL As New Vector3D
Dim startR As New Vector2D
Dim endR As New Vector2D
Dim i As Integer
Dim j As Integer

For i = 1 To NumofNodes
    Nodes(i).Lightness = 0
    nodenormal.x = Nodes(i).Normal.x
    nodenormal.y = Nodes(i).Normal.y
    nodenormal.Z = Nodes(i).Normal.Z
    For j = 1 To Parent.NumberofLights
        Set LightP = VectorMinus(Parent.GetLight(j).LightPoint, Nodes(i))
        lightcosangle = VectorCosAngle(nodenormal, LightP)
        
        tempL.x = Parent.GetLight(j).LightPoint.x
        tempL.y = Parent.GetLight(j).LightPoint.y
        tempL.Z = Parent.GetLight(j).LightPoint.Z
        
        Set MirrorP = VectorMinus(TimesVector(2, TimesVector(DotProduct(nodenormal, tempL), nodenormal)), tempL)
        mirrorcosangle = VectorCosAngle(MirrorP, Parent.ActiveCamera.CameraPoint)
        
        Nodes(i).Lightness = Nodes(i).Lightness + Parent.GetLight(j).Ambient   'AMBIENT LIGHT
        If lightcosangle > 0 Then                           'DIFFUSE LIGHT
            Nodes(i).Lightness = Nodes(i).Lightness + Parent.GetLight(j).Intensity * Kd * lightcosangle
            If mirrorcosangle > 0 Then                          'SPECULAR LIGHT
                Nodes(i).Lightness = Nodes(i).Lightness + Parent.GetLight(j).Intensity * Ks * mirrorcosangle ^ n
            End If
        End If
    Next
    Nodes(i).Lightness = Nodes(i).Lightness + Lightness
Next
End Sub

Public Property Get minX() As Long
Dim minimumX As Long

minX = 2000000
For i = 1 To NumofFaces
    If Faces(i).Hidden = False Then
        minimumX = Faces(i).minX
        If minimumX < minX Then
            minX = minimumX
        End If
    End If
Next
End Property
Public Property Get maxX() As Long
Dim maximumX As Long

maxX = -2000000
For i = 1 To NumofFaces
    If Faces(i).Hidden = False Then
        maximumX = Faces(i).maxX
        If maximumX > maxX Then
            maxX = maximumX
        End If
    End If
Next
End Property
Public Property Get minY() As Long
Dim minimumY As Long

minY = 2000000
For i = 1 To NumofFaces
    If Faces(i).Hidden = False Then
        minimumY = Faces(i).minY
        If minimumY < minY Then
            minY = minimumY
        End If
    End If
Next
End Property
Public Property Get maxY() As Long
Dim maximumY As Long

maxY = -2000000
For i = 1 To NumofFaces
    If Faces(i).Hidden = False Then
        maximumY = Faces(i).maxY
        If maximumY > maxY Then
            maxY = maximumY
        End If
    End If
Next
End Property


