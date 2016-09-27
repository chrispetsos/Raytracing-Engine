Attribute VB_Name = "RayTrace"
'Private Const MAX_DEPTH = 2
'Private Const WeightTreshold = 0.1
Public D3DMeshes() As D3DXMesh
Public FaceHit As Long
Public RayTraceDepth As Integer
Public WeightTreshold As Double

'totally ignores transmission
'avoids internal reflection
'Uses D3D for all calculations!!!

Public Function TraceRay(start As D3DVECTOR, direction As D3DVECTOR, depth As Integer, color As D3DCOLORVALUE, CumulativeWeight As Double) As Integer
Dim IntersectionPoint As D3DVECTOR
Dim IntersectionPointNormal As D3DVECTOR
Dim ReflectedDirection As D3DVECTOR
'Dim TransmittedDirection As Vector3D
Dim LocalColor As D3DCOLORVALUE
Dim ReflectedColor As D3DCOLORVALUE
Dim TransmittedColor As D3DCOLORVALUE
Dim i As Integer
Dim j As Integer
Dim minraydist As Double
Dim tempdist As Double
Dim ObjectHit As Integer
Dim tempinter As D3DVECTOR
Dim tempinternormal As D3DVECTOR
Dim dif As D3DVECTOR
Dim diflen As Double
Dim tempcol As D3DCOLORVALUE
Dim retdist As Single
Dim retFaceIndex As Long
Dim result As Boolean
Dim tempface As D3DXMesh
'FaceHit = 0

If depth > RayTraceDepth Or CumulativeWeight < WeightTreshold Then
    color.r = 0     'Adaptive Depth Control
    color.g = 0
    color.b = 0
Else
    'Intersect ray with all objects and find intersection point (if any)
    'that is closest to ray
'    MainForm.NumOfRays = MainForm.NumOfRays + 1
    TraceRay = 0
    minraydist = 1000000000
    For i = 1 To NumOfMeshes
        If g_Mesh(i).IntersectRayBound(start, direction) = True Then
            'PATCH INTERSECTION     VERY FAST WITH D3D!!!!!
'            For j = 1 To theWorld.GetObject(i).numoffaces
'                TriToMesh theWorld.GetObject(i).getFace(j), tempface
                result = IntersectMesh(g_Mesh(i), start, direction, retdist, tempinter, tempinternormal, retFaceIndex)
                If result Then
    '                If retdist < 0.1 Then Debug.Assert False
    '                If retdist < minRayDist And retdist > 0.1 Then
                    If retdist < minraydist Then
    '                If retdist < minRayDist And retdist > 0.1 And FaceHit <> retFaceIndex Then
                        minraydist = retdist
                        ObjectHit = i               'intersection point
                        IntersectionPoint = tempinter
                        IntersectionPointNormal = tempinternormal
                        FaceHit = retFaceIndex
                        TraceRay = 1
                    Else
                    End If
                End If
'            Next
        End If
    Next
    
    If TraceRay = 0 Then        'no intersection
        tempcol.r = 0
        tempcol.g = 0
        tempcol.b = 0
    Else
    '    localcolor = contribution of local color model at IntersectionPoint
        LocalColor = CalcLocalColor(ObjectHit, IntersectionPoint, IntersectionPointNormal)
    '    Calculate direction of reflected ray and fire it
        result = CalcReflectedDirection(direction, IntersectionPointNormal, ReflectedDirection)
'        result = CalcReflectedDirection(direction, theWorld.GetObject(ObjectHit).getFace(retFaceIndex).GetNormalD3D, ReflectedDirection)
        If result Then
            'Move the intersection point a bit along the reflected direction so
            'that it won't reflect itself
'            D3DXVec3Scale tempinter, IntersectionPointNormal, 0.000001
'            D3DXVec3Scale tempinter, ReflectedDirection, 0.001
'            D3DXVec3Scale tempinter, ReflectedDirection, 0
'            D3DXVec3Add tempinter, IntersectionPoint, tempinter
'            TraceRay theWorld, tempinter, ReflectedDirection, depth + 1, ReflectedColor, CumulativeWeight * theWorld.GetObject(ObjectHit).Krg
            TraceRay IntersectionPoint, ReflectedDirection, depth + 1, ReflectedColor, CumulativeWeight * g_Mesh(ObjectHit).Krg
        End If
'    '    Calculate direction of transmitted ray and fire it
'        Set TransmittedDirection = CalcTransmittedDirection(direction, IntersectionPoint.Normal, theWorld.GetObject(ObjectHit).hta)
'        If Not (TransmittedDirection Is Nothing) Then
'            'Move the intersection point a bit along the transmitted direction so
'            'that it won't refract itself
'            Set tempinter = VectorPlus(IntersectionPoint, TimesVector(0.001, TransmittedDirection))
'            TraceRay theWorld, tempinter, TransmittedDirection, depth + 1, TransmittedColor, CumulativeWeight * theWorld.GetObject(ObjectHit).Ktg
'        End If

'       combine color, LocalColor, LocalWeightForSurface, ReflectedColor,
'       ReflectedWeightForSurface, TransmittedColor,
'       TransmittedWeightForSurface

'       ATTENTION!!!    TRANSMITTED PARAMETER MISSING
        
'        tempcol.r = theWorld.GetObject(ObjectHit).Kl * LocalColor.r + theWorld.GetObject(ObjectHit).Krg * ReflectedColor.r ' + theWorld.GetObject(ObjectHit).Ktg * TransmittedColor.r
'        tempcol.g = theWorld.GetObject(ObjectHit).Kl * LocalColor.g + theWorld.GetObject(ObjectHit).Krg * ReflectedColor.g ' + theWorld.GetObject(ObjectHit).Ktg * TransmittedColor.g
'        tempcol.b = theWorld.GetObject(ObjectHit).Kl * LocalColor.b + theWorld.GetObject(ObjectHit).Krg * ReflectedColor.b ' + theWorld.GetObject(ObjectHit).Ktg * TransmittedColor.b
'        D3DXColorScale tempcol, LocalColor, 1
        D3DXColorScale tempcol, ReflectedColor, g_Mesh(ObjectHit).Krg
'        D3DXColorModulate tempcol, tempcol, LocalColor
        D3DXColorAdd tempcol, tempcol, LocalColor
    End If
End If
color = tempcol
End Function

Public Function CalcReflectedDirection(RayDir As D3DVECTOR, IntersectionPointNormal As D3DVECTOR, ReflectedDirection As D3DVECTOR) As Boolean
Dim Costh As Double
Dim CosthF As Double
Dim CosthR As Double
Dim temp3d As D3DVECTOR

D3DXVec3Scale temp3d, RayDir, -1
Costh = D3DXVec3Dot(IntersectionPointNormal, temp3d)
'Don't let the reflected direction get "into" the object
'CosthF = D3DXVec3Dot(FaceNormal, temp3d)
'If CosthF < Costh Then
'    D3DXVec3Scale temp3d, FaceNormal, 2 * CosthF
'If Costh > 0 Then
    D3DXVec3Scale temp3d, IntersectionPointNormal, 2 * Costh
    D3DXVec3Add ReflectedDirection, RayDir, temp3d
    CalcReflectedDirection = True
'ElseIf CosthF > 0 Then
'    D3DXVec3Scale temp3d, FaceNormal, 2 * CosthF
'    D3DXVec3Add ReflectedDirection, RayDir, temp3d
'    CalcReflectedDirection = True
'Else
'End If
'Else
'    D3DXVec3Scale temp3d, FaceNormal, 2 * CosthF
'    D3DXVec3Add ReflectedDirection, RayDir, temp3d
'    CalcReflectedDirection = True
'End If
'D3DXVec3Normalize ReflectedDirection, ReflectedDirection
'CosthR = D3DXVec3Dot(ReflectedDirection, FaceNormal)
'If CosthR < 0 Then
'    CalcReflectedDirection = False
'End If
End Function

'Public Function CalcTransmittedDirection(RayDir As Vector3D, IntersectionPointNormal As Vector3D, hta As Double) As Vector3D
'Dim Costh1 As Double
'Dim Costh2 As Double
'Dim temp3d As New Vector3D
'Dim root As Double
'
'Costh1 = VectorCosAngle(IntersectionPointNormal, TimesVector(-1, RayDir))
'root = 1 - (1 - Costh1 ^ 2) / (hta ^ 2)
'Costh2 = hta * Sqr(root)
'Set temp3d = VectorMinus(TimesVector(1 / hta, RayDir), TimesVector(Costh2 - Costh1 / hta, IntersectionPointNormal))
'Set CalcTransmittedDirection = temp3d
'End Function

Public Function intersect_RaySphere(StartRay As D3DVECTOR, RayDir As D3DVECTOR, Center As D3DVECTOR, Radius As Double, IntersectionPoint As D3DVECTOR, IntersectionPointNormal As D3DVECTOR, Distance As Double) As Boolean
Dim L As D3DVECTOR
Dim Tca As Double
Dim d2 As Double
Dim Thc As Double
Dim t As Double
Dim temp3d As D3DVECTOR
Dim temp3d2 As D3DVECTOR
Dim len1 As D3DVECTOR
Dim len2 As D3DVECTOR
Dim l1 As Double
Dim l2 As Double

D3DXVec3Subtract L, Center, StartRay
Tca = D3DXVec3Dot(L, RayDir)
If Tca < 0 Then
    Exit Function
End If

d2 = D3DXVec3Dot(L, L) - Tca ^ 2
If d2 > Radius ^ 2 Then
    Exit Function
Else                'Ray intersects Sphere
    Thc = Sqr(Radius ^ 2 - d2)
    t = Tca - Thc
    If t >= 0 Then
        D3DXVec3Scale temp3d, RayDir, t
        D3DXVec3Add temp3d, StartRay, temp3d
    End If
    t = Tca + Thc
    If t >= 0 Then
        D3DXVec3Scale temp3d2, RayDir, t
        D3DXVec3Add temp3d2, StartRay, temp3d2
    End If
    D3DXVec3Subtract len1, temp3d, StartRay
    D3DXVec3Subtract len2, temp3d2, StartRay
    l1 = D3DXVec3Length(len1)
    l2 = D3DXVec3Length(len2)
    If temp3d.X = 0 And temp3d.Y = 0 And temp3d.z = 0 Then
        IntersectionPointNormal.X = (temp3d2.X - Center.X) / Radius
        IntersectionPointNormal.Y = (temp3d2.Y - Center.Y) / Radius
        IntersectionPointNormal.z = (temp3d2.z - Center.z) / Radius
        IntersectionPoint = temp3d2
        Distance = l2
        intersect_RaySphere = True
    ElseIf temp3d2.X = 0 And temp3d2.Y = 0 And temp3d2.z = 0 Then
        IntersectionPointNormal.X = (temp3d.X - Center.X) / Radius
        IntersectionPointNormal.Y = (temp3d.Y - Center.Y) / Radius
        IntersectionPointNormal.z = (temp3d.z - Center.z) / Radius
        IntersectionPoint = temp3d
        Distance = l1
        intersect_RaySphere = True
    ElseIf l1 < l2 Then
        IntersectionPointNormal.X = (temp3d.X - Center.X) / Radius
        IntersectionPointNormal.Y = (temp3d.Y - Center.Y) / Radius
        IntersectionPointNormal.z = (temp3d.z - Center.z) / Radius
        IntersectionPoint = temp3d
        Distance = l1
        intersect_RaySphere = True
    Else
        IntersectionPointNormal.X = (temp3d2.X - Center.X) / Radius
        IntersectionPointNormal.Y = (temp3d2.Y - Center.Y) / Radius
        IntersectionPointNormal.z = (temp3d2.z - Center.z) / Radius
        IntersectionPoint = temp3d2
        Distance = l2
        intersect_RaySphere = True
    End If
End If
End Function

'Public Function intersect_RayBound(StartRay As D3DVECTOR, RayDir As D3DVECTOR) As Boolean
'Dim L As D3DVECTOR
'Dim Tca As Double
'Dim d2 As Double
'
'D3DXVec3Subtract L, theObject.GetD3DBoundCenter, StartRay
'Tca = D3DXVec3Dot(L, RayDir)
'd2 = D3DXVec3Dot(L, L) - Tca ^ 2
'If d2 > theObject.BoundRadius ^ 2 Then
'    intersect_RayBound = False
'    Exit Function
'Else
'    intersect_RayBound = True
'End If
'End Function

'Public Function Point2DWorld(thePoint As Vector2D) As Vector2D
'Dim temp2d As New Vector2D
'
'temp2d.X = 1 * (-MainForm.theWorld.RasterDevice.ScaleHeight / 2 + thePoint.X)
'temp2d.Y = 1 * (MainForm.theWorld.RasterDevice.ScaleWidth / 2 - thePoint.Y)
'
'Set Point2DWorld = temp2d
'End Function
'
'Public Function RayOnViewPlane(theCamera As Camera, thePixel As Vector2D) As Vector3D
'Dim temp3d As New Vector3D
'Dim theWorldPoint As New Vector2D
'Dim temparr1() As Double
'Dim tempPoint() As Double
'Dim tempPoint3D() As Double
'
''Mporw na beltistopoihsw an briskw to World Point xwris na ftiaxnw pinakes
''alla briskontas sto xarti ta stoixeia tou pinaka pou dinoun to shmeio
'Set theWorldPoint = Point2DWorld(thePixel)
'With theCamera
'    PrepareArray temparr1, 4, 4, .iVm11, .iVm12, .iVm13, .iVm14, .iVm21, .iVm22, .iVm23, .iVm24, .iVm31, .iVm32, .iVm33, .iVm34, .iVm41, .iVm42, .iVm43, .iVm44
'End With
'PrepareArray tempPoint, 1, 4, theWorldPoint.X, theWorldPoint.Y, 1000, 1
'MultiplyArray tempPoint, temparr1, tempPoint3D
'temp3d.X = tempPoint3D(1, 1)
'temp3d.Y = tempPoint3D(1, 2)
'temp3d.z = tempPoint3D(1, 3)
'
'Set RayOnViewPlane = temp3d
'End Function

Public Function CalcLocalColor(ObjectIntersected As Integer, IntersectionPoint As D3DVECTOR, IntersectionPointNormal As D3DVECTOR) As D3DCOLORVALUE
Dim i As Integer
Dim LightP As D3DVECTOR
Dim Lightcosangle As Double
Dim MirrorP As D3DVECTOR
Dim Mirrorcosangle As Double
Dim TransP As D3DVECTOR
Dim Transcosangle As Double
Dim tempcol As D3DCOLORVALUE
Dim tempDif As D3DCOLORVALUE
Dim tempSpecI As D3DCOLORVALUE
Dim tempSpec As D3DCOLORVALUE
Dim tempTran As D3DCOLORVALUE
Dim tempAmb As D3DCOLORVALUE
Dim Ip As Double
Dim cR As Single
Dim cG As Single
Dim cB As Single
Dim temp3d As D3DVECTOR
Dim CurMat As D3DMATERIAL8
Dim ambCol As Long

CurMat = g_Mesh(ObjectIntersected).GetMaterial(0)
Ip = 0
'tempSpec = CurMat.specular
'D3DXVec3Normalize IntersectionPointNormal, IntersectionPointNormal
ambCol = g_dev.GetRenderState(D3DRS_AMBIENT)
For i = 1 To NumOfLights
    If Lights(i).GetState = 1 Then
        D3DXVec3Subtract LightP, Lights(i).GetD3DLight.Position, IntersectionPoint
        D3DXVec3Normalize LightP, LightP
        Lightcosangle = D3DXVec3Dot(IntersectionPointNormal, LightP)
        
        D3DXVec3Scale temp3d, IntersectionPointNormal, 2 * Lightcosangle
        D3DXVec3Subtract MirrorP, temp3d, LightP
        D3DXVec3Normalize MirrorP, MirrorP
        temp3d = Cameras(curCamera).GetCameraPoint
        D3DXVec3Normalize temp3d, temp3d
        Mirrorcosangle = D3DXVec3Dot(MirrorP, temp3d)
    '    Set TransP = TimesVector(-1 / (theWorld.GetObject(ObjectIntersected).hta - 0.999999), VectorMinus(TimesVector(-1, LightP), TimesVector(theWorld.GetObject(ObjectIntersected).hta, theWorld.ActiveCamera.CameraPoint)))
    '    normalize TransP
    '    Transcosangle = VectorCosAngle(TransP, IntersectionPoint.Normal)
        If Lightcosangle > 0 Then  '0.001 for floating point errors
            'DIFFUSE LIGHT
    '        Ip = Ip + theWorld.GetLight(i).Intensity * theWorld.GetObject(ObjectIntersected).Kd * Lightcosangle
            D3DXColorModulate tempDif, Lights(i).GetD3DLight.diffuse, CurMat.diffuse
    '        D3DXColorScale tempDif, tempDif, g_Mesh(ObjectIntersected).Kd * Lightcosangle
            D3DXColorScale tempDif, tempDif, Lightcosangle
            D3DXColorAdd tempcol, tempcol, tempDif
            If Mirrorcosangle > 0 Then
    '            'SPECULAR LIGHT
    '            Ip = Ip + theWorld.GetLight(i).Intensity * theWorld.GetObject(ObjectIntersected).Ks * Mirrorcosangle ^ theWorld.GetObject(ObjectIntersected).n
    '            D3DXColorScale tempSpec, tempSpec, g_Mesh(ObjectIntersected).Ks * Mirrorcosangle ^ g_Mesh(ObjectIntersected).n
                D3DXColorScale tempSpecI, Lights(i).GetD3DLight.specular, Mirrorcosangle ^ CurMat.power
                D3DXColorAdd tempSpec, tempSpec, tempSpecI
                D3DXColorModulate tempSpec, tempSpec, CurMat.specular
                D3DXColorAdd tempcol, tempcol, tempSpec
'                D3DXColorAdd tempcol, tempcol, tempSpec
            End If
    ''        'TRANSMITTED LIGHT
    ''        Ip = Ip + theWorld.GetLight(i).Intensity * theWorld.GetObject(ObjectIntersected).Ktg * Transcosangle ^ theWorld.GetObject(ObjectIntersected).n
            If Lights(i).CastShadows = 1 Then
    '            Ip = Ip * ShadowAttenuation(theWorld, theWorld.GetLight(i), IntersectionPoint, IntersectionPointNormal)
                D3DXColorScale tempcol, tempcol, ShadowAttenuation(Lights(i), IntersectionPoint, IntersectionPointNormal)
            End If
        End If
    '    'AMBIENT LIGHT
    '    Ip = Ip + theWorld.GetLight(i).Ambient
    '    D3DXColorAdd tempcol, tempcol, CurMat.Ambient
        D3DXColorAdd tempAmb, tempAmb, Lights(i).GetD3DLight.Ambient  'sum(Lai)
    '    D3DXColorScale tempAmb, tempAmb, 1
    End If
Next

'D3DXColorModulate tempSpec, tempSpec, CurMat.specular
'D3DXColorAdd tempcol, tempcol, tempSpec
'For i = 1 To NumOfLights
'    If Lights(i).CastShadows = 1 Then
'        D3DXColorScale tempcol, tempcol, ShadowAttenuation(Lights(i), IntersectionPoint, IntersectionPointNormal)
'    End If
'Next
D3DXColorAdd tempAmb, tempAmb, LONGtoD3DCOLORVALUE(ambCol)  'Ga
D3DXColorModulate tempAmb, tempAmb, CurMat.Ambient          'Mc
D3DXColorAdd tempcol, tempcol, tempAmb
'Ip = Ip + theWorld.GetObject(ObjectIntersected).Lightness
'If Ip < 0 Then Ip = 0
'If Ip > 1 Then Ip = 1
'hls2rgb theWorld.GetObject(ObjectIntersected).hue, Ip, theWorld.GetObject(ObjectIntersected).Saturation, cR, cG, cB
'tempcol.r = cR
'tempcol.g = cG
'tempcol.b = cB
CalcLocalColor = tempcol

End Function

Public Function ShadowAttenuation(theLight As Light, IntersectionPoint As D3DVECTOR, IntersectionPointNormal As D3DVECTOR) As Double
Dim i As Integer
Dim j As Integer
Dim start As D3DVECTOR
Dim direction As D3DVECTOR
Dim InterP As D3DVECTOR
Dim tempinter As D3DVECTOR
Dim tempinternormal As D3DVECTOR
Dim tempdist As Double
Dim retdist As Single
Dim retFaceIndex As Long

'Intersect light ray with all objects and if you find intersection
'intersection point is in the shadow of that light
'Attenuate it by the Ktg factor of the object intersected
ShadowAttenuation = 1
minraydist = 1000000000
D3DXVec3Subtract direction, theLight.GetD3DLight.Position, IntersectionPoint
D3DXVec3Normalize direction, direction
'move intersection point a bit so that it won't shadow itself
'first "out" along its normal
D3DXVec3Add InterP, IntersectionPoint, IntersectionPointNormal
'second "towards" to light
'Set interP = VectorPlus(interP, TimesVector(1, direction))
start = InterP
For i = 1 To NumOfMeshes
    If g_Mesh(i).IntersectRayBound(start, direction) = True Then
    'PATCH INTERSECTION     VERY FAST WITH D3D!!!!!
'            For j = 1 To theWorld.GetObject(i).numoffaces
'                TriToMesh theWorld.GetObject(i).getFace(j), tempface
        result = IntersectMesh(g_Mesh(i), start, direction, retdist, tempinter, tempinternormal, retFaceIndex)
        If result Then
            ShadowAttenuation = ShadowAttenuation * g_Mesh(i).Ktg
            If ShadowAttenuation = 0 Then Exit Function
    '                If retdist < 0.1 Then Debug.Assert False
    '                If retdist < minRayDist And retdist > 0.1 Then
'            If retdist < minraydist Then
    '                If retdist < minRayDist And retdist > 0.1 And FaceHit <> retFaceIndex Then
'                minraydist = retdist
'                ObjectHit = i               'intersection point
'                IntersectionPoint = tempinter
'                IntersectionPointNormal = tempinternormal
'                FaceHit = retFaceIndex
'                TraceRay = 1
'            Else
'            End If
        End If
'            Next
    End If
Next


'For i = 1 To theWorld.NumberOfObjects
'    If intersect_RayBound(start, direction, theWorld.GetObject(i)) = True Then
'        If theWorld.GetObject(i).Representation = SPHERE Then
'            'SPHERE INTERSECTION    VERY FAST!!!
'            result = intersect_RaySphere(start, direction, theWorld.GetObject(i).GetD3DBoundCenter, theWorld.GetObject(i).BoundRadius - 5, tempinter, tempinternormal, tempdist)
'            If result Then
'                ShadowAttenuation = ShadowAttenuation * theWorld.GetObject(i).Ktg
'                If ShadowAttenuation = 0 Then Exit Function
'            End If
'        ElseIf theWorld.GetObject(i).Representation = BOX Then
'            'BOX INTERSECTION    VERY FAST WITH D3D!!!
'            result = IntersectBox(D3DMeshes(i), theWorld.GetObject(i), start, direction, retdist, tempinter, tempinternormal)
'            If result Then
'                ShadowAttenuation = ShadowAttenuation * theWorld.GetObject(i).Ktg
'                If ShadowAttenuation = 0 Then Exit Function
'            End If
'        ElseIf theWorld.GetObject(i).Representation = PATCH Then
'            'PATCH INTERSECTION     VERY FAST WITH D3D!!!!!
'            result = IntersectMesh(D3DMeshes(i), theWorld.GetObject(i), start, direction, retdist, tempinter, tempinternormal, retFaceIndex)
'            If result Then
'                ShadowAttenuation = ShadowAttenuation * theWorld.GetObject(i).Ktg
'                If ShadowAttenuation = 0 Then Exit Function
'            End If
'        End If
'    End If
'Next
End Function
