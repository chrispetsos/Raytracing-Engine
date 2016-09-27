Attribute VB_Name = "ScanFill"
Type Mpoint
    NodeNo As Integer
    x As Double
    y As Double
End Type

Public Sub doGouraud(RasterDevice As PictureBox, pts() As Mpoint, Intensity() As Double, theFace As Face)
Dim upvertex As Integer             'to PANW shmeio tou trigwnou
Dim middlevertex As Integer         'to MESAIO shmeio tou trigwnou
Dim lowvertex As Integer            'to KATW shmeio tou trigwnou
Dim Minimun 'proswrnh metablhth sugkrishs
Dim Maximum 'proswrnh metablhth sugkrishs
Dim a(1 To 3) As Double 'to a ths eutheias
Dim b(1 To 3) As Double 'to b ths eutheias
Dim I1 As Double    'h FWTEINOTHTA ths prwths koryfhs
Dim I2 As Double    'h FWTEINOTHTA ths deuterhs koryfhs
Dim I3 As Double    'h FWTEINOTHTA ths triths koryfhs
Dim Ia As Double    'h FWTEINOTHTA ths ARISTERHS pleyras
Dim Ib As Double    'h FWTEINOTHTA ths DEKSIAS pleyras
Dim Ip As Double    'h FWTEINOTHTA tou ESWTERIKOU shmeiou
Dim Ys As Integer   'syntetagmenh Y ths orizontias scanline eutheias
Dim Xs As Integer   'syntetagmenh X ths orizontias scanline eutheias
Dim left As Integer     'h ARISTERH pleura pou temnei h scanline
Dim right As Integer    'h DEKSIA pleura pou temnei h scanline
Dim Xa As Integer   'syntetagmenh X ths scanline me thn ARISTERH pleura
Dim Xb As Integer   'syntetagmenh X ths scanline me thn DEKSIA pleura
Dim cR As Single
Dim cG As Single
Dim cB As Single
Dim incol As Long   'xrwma tou shmeiou pou tha baftei

'Parametroi gia incremental
Dim DIp As Double
Dim Dx As Double
Dim Ipn As Double

Minimun = 200000
Maximum = -200000
For i = 1 To 3
    If pts(i).y <= Minimun Then                 'bres thn panw
        Minimun = pts(i).y                      'kai katw korufh
        lowvertex = i
    End If
    If pts(i).y > Maximum Then
        Maximum = pts(i).y
        upvertex = i
    End If
Next
If upvertex = 1 And lowvertex = 2 Then          'twra bres kai thn
    middlevertex = 3                            'mesaia korufh
ElseIf upvertex = 1 And lowvertex = 3 Then
    middlevertex = 2
ElseIf upvertex = 2 And lowvertex = 1 Then
    middlevertex = 3
ElseIf upvertex = 2 And lowvertex = 3 Then
    middlevertex = 1
ElseIf upvertex = 3 And lowvertex = 1 Then
    middlevertex = 2
ElseIf upvertex = 3 And lowvertex = 2 Then
    middlevertex = 1
Else
    middlevertex = 1
End If

For i = 1 To 3                                  'upologise ta a,b
    k1 = i
    k2 = i + 1
    If k2 = 4 Then k2 = 1
    If pts(k1).x - pts(k2).x = 0 Then
        'katakorufh pleura ths morfhs x=b kai shmeia (b,y)
        a(i) = 1000000          'APEIRO
        b(i) = pts(k1).x        'x=b
    Else
        a(i) = (pts(k1).y - pts(k2).y) / (pts(k1).x - pts(k2).x)
        b(i) = pts(k1).y - a(i) * pts(k1).x
    End If
Next

'orise thn aristerh kai deksia pleura
If upvertex = 1 Then
    left = 3
    right = 1
ElseIf upvertex = 2 Then
    left = 1
    right = 2
ElseIf upvertex = 3 Then
    left = 2
    right = 3
End If

Y1 = pts(upvertex).y
Y2 = pts(OtherVertice(left, upvertex)).y
Y3 = pts(OtherVertice(right, upvertex)).y
I1 = Intensity(upvertex)                    'metablhtes fwteinothtas
I2 = Intensity(OtherVertice(left, upvertex))
I3 = Intensity(OtherVertice(right, upvertex))
For Ys = pts(upvertex).y To pts(middlevertex).y Step -1 'gia ta Y apo thn PANW
                                                        'ws th MESAIA korufh
    If Y1 = Y2 Then
        'orizontia aristerh ths upvertex pleura
        Ia = I2
    Else
        Ia = (I1 * (Ys - Y2) + I2 * (Y1 - Ys)) / (Y1 - Y2)
    End If
    If Y1 = Y3 Then
        'orizontia deksia ths upvertex pleura
        Ib = I3
    Else
        Ib = (I1 * (Ys - Y3) + I3 * (Y1 - Ys)) / (Y1 - Y3)
    End If
    If a(left) = 0 Then
        'orizontia pleura
        Xa = Min(pts(middlevertex).x, pts(upvertex).x)
    ElseIf a(left) = 1000000 Then
        'katakorufh pleura
        Xa = b(left)
    Else
        Xa = (Ys - b(left)) / a(left)
    End If
    If a(right) = 0 Then
        'orizontia pleura
        Xb = Max(pts(middlevertex).x, pts(upvertex).x)
    ElseIf a(right) = 1000000 Then
        'katakorufh pleura
        Xb = b(right)
    Else
        Xb = (Ys - b(right)) / a(right)
    End If
    Ip = Ia
    Dx = 1
    If Xb = Xa Then
        DIp = Ia
    Else
        DIp = Dx * (Ib - Ia) / (Xb - Xa)
    End If
    For Xs = Xa To Xb       'grapse ta X ths scanline
        If Xa = Xb Then
            'katakorufh pleura
            Ip = Ia
        Else
            Ip = Ip + DIp
        End If
        If Ip < 0 Then Ip = 0
        If Ip > 1 Then Ip = 1
        hls2rgb theFace.Parent.hue, Ip, theFace.Parent.Saturation, cR, cG, cB
'        cR = theColor.r * Ip
'        cG = theColor.g * Ip
'        cB = theColor.b * Ip
        incol = RGB(cR * 255, cG * 255, cB * 255)
        PSetAPI RasterDevice, Xs, Ys, incol
    Next
Next

'orise thn aristerh kai deksia pleura
If lowvertex = 1 Then
    left = 1
    right = 3
ElseIf lowvertex = 2 Then
    left = 2
    right = 1
ElseIf lowvertex = 3 Then
    left = 3
    right = 2
End If

Y1 = pts(lowvertex).y
Y2 = pts(OtherVertice(left, lowvertex)).y
Y3 = pts(OtherVertice(right, lowvertex)).y
I1 = Intensity(lowvertex)                    'metablhtes fwteinothtas
I2 = Intensity(OtherVertice(left, lowvertex))
I3 = Intensity(OtherVertice(right, lowvertex))
For Ys = pts(middlevertex).y To pts(lowvertex).y Step -1    'gia ta Y apo thn MESAIA
                                                            'ws th KATW korufh
    If Y1 = Y2 Then
        'orizontia aristerh ths lowvertex pleura
        Ia = I2
    Else
        Ia = (I1 * (Y2 - Ys) + I2 * (Ys - Y1)) / (Y2 - Y1)
'        Ia = (I1 * (Ys - Y2) + I2 * (Y1 - Ys)) / (Y1 - Y2)
    End If
    If Y1 = Y3 Then
        'orizontia deksia ths lowvertex pleura
        Ib = I3
    Else
        Ib = (I1 * (Ys - Y3) + I3 * (Y1 - Ys)) / (Y1 - Y3)
    End If
    If a(left) = 0 Then
        'orizontia pleura
        Xa = Min(pts(middlevertex).x, pts(lowvertex).x)
    ElseIf a(left) = 1000000 Then
        'katakorufh pleura
        Xa = b(left)
    Else
        Xa = (Ys - b(left)) / a(left)
    End If
    If a(right) = 0 Then
        'orizontia pleura
        Xb = Max(pts(middlevertex).x, pts(lowvertex).x)
    ElseIf a(right) = 1000000 Then
        'katakorufh pleura
        Xb = b(right)
    Else
        Xb = (Ys - b(right)) / a(right)
    End If
    Ip = Ia
    Dx = 1
    If Xb = Xa Then
        DIp = Ia
    Else
        DIp = Dx * (Ib - Ia) / (Xb - Xa)
    End If
    For Xs = Xa To Xb       'grapse ta X ths scanline
        If Xa = Xb Then
            'katakorufh pleura
            Ip = Ia
        Else
            Ip = Ip + DIp
        End If
        If Ip < 0 Then Ip = 0
        If Ip > 1 Then Ip = 1
        hls2rgb theFace.Parent.hue, Ip, theFace.Parent.Saturation, cR, cG, cB
'        cR = theColor.r * Ip
'        cG = theColor.g * Ip
'        cB = theColor.b * Ip
        incol = RGB(cR * 255, cG * 255, cB * 255)
        PSetAPI RasterDevice, Xs, Ys, incol
    Next
Next
End Sub

Public Sub doPhong(theWorld As World, pts() As Mpoint, Nodes() As Vector3D, Normals() As Vector3D, theColor As FaceColor)
Dim upvertex As Integer             'to PANW shmeio tou trigwnou
Dim middlevertex As Integer         'to MESAIO shmeio tou trigwnou
Dim lowvertex As Integer            'to KATW shmeio tou trigwnou
Dim Minimun 'proswrnh metablhth sugkrishs
Dim Maximum 'proswrnh metablhth sugkrishs
Dim a(1 To 3) As Double 'to a ths eutheias
Dim b(1 To 3) As Double 'to b ths eutheias
Dim Node1 As New Vector3D      'h prwth korufh
Dim Node2 As New Vector3D      'h deuterh korufh
Dim Node3 As New Vector3D      'h trith korufh
Dim Nodea As New Vector3D      'h oura tou Normal sthn ARISTERH pleura
Dim Nodeb As New Vector3D      'h oura tou Normal sthn DEKSIA pleura
Dim Nodep As New Vector3D      'h oura tou Normal sto ESWTERIKO shmeio
Dim N1 As New Vector3D      'to Normal ths prwths korufhs
Dim N2 As New Vector3D      'to Normal ths deuterhs korufhs
Dim N3 As New Vector3D      'to Normal ths triths korufhs
Dim Na As New Vector3D      'to Normal sthn ARISTERH pleura
Dim Nb As New Vector3D      'to Normal sthn DEKSIA pleura
Dim Np As New Vector3D      'to Normal sto ESWTERIKO shmeio
Dim LightP As New Vector3D  'to vector tou fwtos
Dim tempL As New Vector3D  'to vector tou fwtos
Dim lightcosangle As Double ' h gwnia tou Np me to LightP
Dim MirrorP As New Vector3D     'to Mirror Vector gia to Specular
Dim mirrorcosangle As Double    'h gwnia tou MirrorP me thn Camera
Dim tempV As New Vector3D
Dim Ip As Double    'h FWTEINOTHTA tou eswterikou shmeiou
Dim Ys As Integer   'syntetagmenh Y ths orizontias scanline eutheias
Dim Xs As Integer   'syntetagmenh X ths orizontias scanline eutheias
Dim left As Integer     'h ARISTERH pleura pou temnei h scanline
Dim right As Integer    'h DEKSIA pleura pou temnei h scanline
Dim Xa As Integer   'syntetagmenh X ths scanline me thn ARISTERH pleura
Dim Xb As Integer   'syntetagmenh X ths scanline me thn DEKSIA pleura
Dim cR As Single
Dim cG As Single
Dim cB As Single
Dim incol As Long   'xrwma tou shmeiou pou tha baftei
Dim j As Integer
Dim q As Integer
Dim w As Integer
Dim IsIn As Boolean

'Parametroi gia incremental
Dim DNp As New Vector3D
Dim DNodep As New Vector3D
Dim Dx As Double

Dx = 1

Minimun = 200000
Maximum = -200000
For i = 1 To 3
    If pts(i).y <= Minimun Then                 'bres thn panw
        Minimun = pts(i).y                      'kai katw korufh
        lowvertex = i
    End If
    If pts(i).y > Maximum Then
        Maximum = pts(i).y
        upvertex = i
    End If
Next
If upvertex = 1 And lowvertex = 2 Then          'twra bres kai thn
    middlevertex = 3                            'mesaia korufh
ElseIf upvertex = 1 And lowvertex = 3 Then
    middlevertex = 2
ElseIf upvertex = 2 And lowvertex = 1 Then
    middlevertex = 3
ElseIf upvertex = 2 And lowvertex = 3 Then
    middlevertex = 1
ElseIf upvertex = 3 And lowvertex = 1 Then
    middlevertex = 2
ElseIf upvertex = 3 And lowvertex = 2 Then
    middlevertex = 1
Else
    middlevertex = 1
End If

For i = 1 To 3                                  'upologise ta a,b
    k1 = i
    k2 = i + 1
    If k2 = 4 Then k2 = 1
    If pts(k1).x - pts(k2).x = 0 Then
        'katakorufh pleura ths morfhs x=b kai shmeia (b,y)
        a(i) = 1000000          'APEIRO
        b(i) = pts(k1).x        'x=b
    Else
        a(i) = (pts(k1).y - pts(k2).y) / (pts(k1).x - pts(k2).x)
        b(i) = pts(k1).y - a(i) * pts(k1).x
    End If
Next

'orise thn aristerh kai deksia pleura
If upvertex = 1 Then
    left = 3
    right = 1
ElseIf upvertex = 2 Then
    left = 1
    right = 2
ElseIf upvertex = 3 Then
    left = 2
    right = 3
End If

Y1 = pts(upvertex).y
Y2 = pts(OtherVertice(left, upvertex)).y
Y3 = pts(OtherVertice(right, upvertex)).y
Set N1 = Normals(upvertex)
Set N2 = Normals(OtherVertice(left, upvertex))
Set N3 = Normals(OtherVertice(right, upvertex))
Set Node1 = Nodes(upvertex)
Set Node2 = Nodes(OtherVertice(left, upvertex))
Set Node3 = Nodes(OtherVertice(right, upvertex))
For Ys = pts(upvertex).y To pts(middlevertex).y Step -1 'gia ta Y apo thn PANW
                                                        'ws th MESAIA korufh
    If Y1 = Y2 Then
        'orizontia aristerh ths upvertex pleura
        Na.x = N2.x
        Na.y = N2.y
        Na.Z = N2.Z
        Nodea.x = Node2.x
        Nodea.y = Node2.y
        Nodea.Z = Node2.Z
    Else
        Na.x = (N1.x * (Ys - Y2) + N2.x * (Y1 - Ys)) / (Y1 - Y2)
        Na.y = (N1.y * (Ys - Y2) + N2.y * (Y1 - Ys)) / (Y1 - Y2)
        Na.Z = (N1.Z * (Ys - Y2) + N2.Z * (Y1 - Ys)) / (Y1 - Y2)
        Nodea.x = (Node1.x * (Ys - Y2) + Node2.x * (Y1 - Ys)) / (Y1 - Y2)
        Nodea.y = (Node1.y * (Ys - Y2) + Node2.y * (Y1 - Ys)) / (Y1 - Y2)
        Nodea.Z = (Node1.Z * (Ys - Y2) + Node2.Z * (Y1 - Ys)) / (Y1 - Y2)
    End If
    If Y1 = Y3 Then
        'orizontia deksia ths upvertex pleura
        Nb.x = N3.x
        Nb.y = N3.y
        Nb.Z = N3.Z
        Nodeb.x = Node3.x
        Nodeb.y = Node3.y
        Nodeb.Z = Node3.Z
    Else
        Nb.x = (N1.x * (Ys - Y3) + N3.x * (Y1 - Ys)) / (Y1 - Y3)
        Nb.y = (N1.y * (Ys - Y3) + N3.y * (Y1 - Ys)) / (Y1 - Y3)
        Nb.Z = (N1.Z * (Ys - Y3) + N3.Z * (Y1 - Ys)) / (Y1 - Y3)
        Nodeb.x = (Node1.x * (Ys - Y3) + Node3.x * (Y1 - Ys)) / (Y1 - Y3)
        Nodeb.y = (Node1.y * (Ys - Y3) + Node3.y * (Y1 - Ys)) / (Y1 - Y3)
        Nodeb.Z = (Node1.Z * (Ys - Y3) + Node3.Z * (Y1 - Ys)) / (Y1 - Y3)
    End If
    If a(left) = 0 Then
        'orizontia pleura
        Xa = Min(pts(middlevertex).x, pts(upvertex).x)
    ElseIf a(left) = 1000000 Then
        'katakorufh pleura
        Xa = b(left)
    Else
        Xa = (Ys - b(left)) / a(left)
    End If
    If a(right) = 0 Then
        'orizontia pleura
        Xb = Max(pts(middlevertex).x, pts(upvertex).x)
    ElseIf a(right) = 1000000 Then
        'katakorufh pleura
        Xb = b(right)
    Else
        Xb = (Ys - b(right)) / a(right)
    End If
'    Dx = Int((3 + Abs(Xa - Xb)) / 3 + 0.5)
    Np.x = Na.x
    Np.y = Na.y
    Np.Z = Na.Z
    Nodep.x = Nodea.x
    Nodep.y = Nodea.y
    Nodep.Z = Nodea.Z
    If Xb = Xa Then
        DNp.x = Na.x
        DNp.y = Na.y
        DNp.Z = Na.Z
        DNodep.x = Nodea.x
        DNodep.y = Nodea.y
        DNodep.Z = Nodea.Z
    Else
        DNp.x = Dx * (Nb.x - Na.x) / (Xb - Xa)
        DNp.y = Dx * (Nb.y - Na.y) / (Xb - Xa)
        DNp.Z = Dx * (Nb.Z - Na.Z) / (Xb - Xa)
        DNodep.x = Dx * (Nodeb.x - Nodea.x) / (Xb - Xa)
        DNodep.y = Dx * (Nodeb.y - Nodea.y) / (Xb - Xa)
        DNodep.Z = Dx * (Nodeb.Z - Nodea.Z) / (Xb - Xa)
    End If
    For Xs = Xa To Xb Step Dx      'grapse ta X ths scanline
        For i = 0 To Dx - 1
            If Xa = Xb Then
                'katakorufh pleura
                Np.x = Na.x
                Np.y = Na.y
                Np.Z = Na.Z
                Nodep.x = Nodea.x
                Nodep.y = Nodea.y
                Nodep.Z = Nodea.Z
            Else
                Np.x = Np.x + DNp.x
                Np.y = Np.y + DNp.y
                Np.Z = Np.Z + DNp.Z
                Nodep.x = Nodep.x + DNodep.x
                Nodep.y = Nodep.y + DNodep.y
                Nodep.Z = Nodep.Z + DNodep.Z
            End If
            Ip = 0
            For j = 1 To theWorld.NumberofLights
                Set LightP = VectorMinus(theWorld.GetLight(j).LightPoint, Nodep)
                lightcosangle = VectorCosAngle(Np, LightP)
                tempL.x = theWorld.GetLight(j).LightPoint.x
                tempL.y = theWorld.GetLight(j).LightPoint.y
                tempL.Z = theWorld.GetLight(j).LightPoint.Z
'                IsIn = False
'                For q = 1 To theWorld.NumberOfObjects  'SHADOW TEST
'                    For w = 1 To theWorld.GetObject(q).NumofFaces
'                        IsIn = IsInShadow(Nodep, theWorld.GetObject(q).getFace(w).getNode(1), theWorld.GetObject(q).getFace(w).getNode(2), theWorld.GetObject(q).getFace(w).getNode(3), tempL)
'                        If IsIn Then Exit For
'                    Next
'                    If IsIn Then Exit For
'                Next
'                If IsIn = False Then
                    Set MirrorP = VectorMinus(TimesVector(2, TimesVector(DotProduct(Np, tempL), Np)), tempL)
                    mirrorcosangle = VectorCosAngle(MirrorP, theWorld.ActiveCamera.CameraPoint)
            
                    Ip = Ip + theWorld.GetLight(j).Ambient  'AMBIENT LIGHT
                    If lightcosangle > 0 Then                           'DIFFUSE LIGHT
                        Ip = Ip + theWorld.GetLight(j).Intensity * Nodes(1).Parent.Kd * lightcosangle
                        If mirrorcosangle > 0 Then                          'SPECULAR LIGHT
                            Ip = Ip + theWorld.GetLight(j).Intensity * Nodes(1).Parent.Ks * mirrorcosangle ^ Nodes(1).Parent.n
                        End If
                    End If
'                Else
'                    Ip = Ip + theWorld.GetLight(j).Ambient  'AMBIENT LIGHT
'                End If
            Next
            Ip = Ip + Nodes(1).Parent.Lightness
            If Ip < 0 Then Ip = 0
            If Ip > 1 Then Ip = 1
            hls2rgb Nodes(1).Parent.hue, Ip, Nodes(1).Parent.Saturation, cR, cG, cB
'            cR = theColor.r * Ip
'            cG = theColor.g * Ip
'            cB = theColor.b * Ip
            incol = RGB(cR * 255, cG * 255, cB * 255)
            PSetAPI theWorld.RasterDevice, Xs + i, Ys, incol
        Next
    Next
Next

'orise thn aristerh kai deksia pleura
If lowvertex = 1 Then
    left = 1
    right = 3
ElseIf lowvertex = 2 Then
    left = 2
    right = 1
ElseIf lowvertex = 3 Then
    left = 3
    right = 2
End If

Y1 = pts(lowvertex).y
Y2 = pts(OtherVertice(left, lowvertex)).y
Y3 = pts(OtherVertice(right, lowvertex)).y
Set N1 = Normals(lowvertex)
Set N2 = Normals(OtherVertice(left, lowvertex))
Set N3 = Normals(OtherVertice(right, lowvertex))
Set Node1 = Nodes(lowvertex)
Set Node2 = Nodes(OtherVertice(left, lowvertex))
Set Node3 = Nodes(OtherVertice(right, lowvertex))
For Ys = pts(middlevertex).y To pts(lowvertex).y Step -1    'gia ta Y apo thn MESAIA
                                                            'ws th KATW korufh
    If Y1 = Y2 Then
        'orizontia aristerh ths upvertex pleura
        Na.x = N2.x
        Na.y = N2.y
        Na.Z = N2.Z
        Nodea.x = Node2.x
        Nodea.y = Node2.y
        Nodea.Z = Node2.Z
    Else
        Na.x = (N1.x * (Ys - Y2) + N2.x * (Y1 - Ys)) / (Y1 - Y2)
        Na.y = (N1.y * (Ys - Y2) + N2.y * (Y1 - Ys)) / (Y1 - Y2)
        Na.Z = (N1.Z * (Ys - Y2) + N2.Z * (Y1 - Ys)) / (Y1 - Y2)
        Nodea.x = (Node1.x * (Ys - Y2) + Node2.x * (Y1 - Ys)) / (Y1 - Y2)
        Nodea.y = (Node1.y * (Ys - Y2) + Node2.y * (Y1 - Ys)) / (Y1 - Y2)
        Nodea.Z = (Node1.Z * (Ys - Y2) + Node2.Z * (Y1 - Ys)) / (Y1 - Y2)
    End If
    If Y1 = Y3 Then
        'orizontia deksia ths upvertex pleura
        Nb.x = N3.x
        Nb.y = N3.y
        Nb.Z = N3.Z
        Nodeb.x = Node3.x
        Nodeb.y = Node3.y
        Nodeb.Z = Node3.Z
    Else
        Nb.x = (N1.x * (Ys - Y3) + N3.x * (Y1 - Ys)) / (Y1 - Y3)
        Nb.y = (N1.y * (Ys - Y3) + N3.y * (Y1 - Ys)) / (Y1 - Y3)
        Nb.Z = (N1.Z * (Ys - Y3) + N3.Z * (Y1 - Ys)) / (Y1 - Y3)
        Nodeb.x = (Node1.x * (Ys - Y3) + Node3.x * (Y1 - Ys)) / (Y1 - Y3)
        Nodeb.y = (Node1.y * (Ys - Y3) + Node3.y * (Y1 - Ys)) / (Y1 - Y3)
        Nodeb.Z = (Node1.Z * (Ys - Y3) + Node3.Z * (Y1 - Ys)) / (Y1 - Y3)
    End If
    If a(left) = 0 Then
        'orizontia pleura
        Xa = Min(pts(middlevertex).x, pts(lowvertex).x)
    ElseIf a(left) = 1000000 Then
        'katakorufh pleura
        Xa = b(left)
    Else
        Xa = (Ys - b(left)) / a(left)
    End If
    If a(right) = 0 Then
        'orizontia pleura
        Xb = Max(pts(middlevertex).x, pts(lowvertex).x)
    ElseIf a(right) = 1000000 Then
        'katakorufh pleura
        Xb = b(right)
    Else
        Xb = (Ys - b(right)) / a(right)
    End If
'    Dx = Int(0.5 + (Abs(Xa - Xb)) / 3)
    Np.x = Na.x
    Np.y = Na.y
    Np.Z = Na.Z
    Nodep.x = Nodea.x
    Nodep.y = Nodea.y
    Nodep.Z = Nodea.Z
    If Xb = Xa Then
        DNp.x = Na.x
        DNp.y = Na.y
        DNp.Z = Na.Z
        DNodep.x = Nodea.x
        DNodep.y = Nodea.y
        DNodep.Z = Nodea.Z
    Else
        DNp.x = Dx * (Nb.x - Na.x) / (Xb - Xa)
        DNp.y = Dx * (Nb.y - Na.y) / (Xb - Xa)
        DNp.Z = Dx * (Nb.Z - Na.Z) / (Xb - Xa)
        DNodep.x = Dx * (Nodeb.x - Nodea.x) / (Xb - Xa)
        DNodep.y = Dx * (Nodeb.y - Nodea.y) / (Xb - Xa)
        DNodep.Z = Dx * (Nodeb.Z - Nodea.Z) / (Xb - Xa)
    End If
    For Xs = Xa To Xb Step Dx      'grapse ta X ths scanline
        For i = 1 To Dx
            If Xa = Xb Then
                'katakorufh pleura
                Np.x = Na.x
                Np.y = Na.y
                Np.Z = Na.Z
                Nodep.x = Nodea.x
                Nodep.y = Nodea.y
                Nodep.Z = Nodea.Z
            Else
                Np.x = Np.x + DNp.x
                Np.y = Np.y + DNp.y
                Np.Z = Np.Z + DNp.Z
                Nodep.x = Nodep.x + DNodep.x
                Nodep.y = Nodep.y + DNodep.y
                Nodep.Z = Nodep.Z + DNodep.Z
            End If
            Ip = 0
            For j = 1 To theWorld.NumberofLights
                Set LightP = VectorMinus(theWorld.GetLight(j).LightPoint, Nodep)
                lightcosangle = VectorCosAngle(Np, LightP)
                tempL.x = theWorld.GetLight(j).LightPoint.x
                tempL.y = theWorld.GetLight(j).LightPoint.y
                tempL.Z = theWorld.GetLight(j).LightPoint.Z
'                IsIn = False
'                For q = 1 To theWorld.NumberOfObjects   'SHADOW TEST
'                    For w = 1 To theWorld.GetObject(q).NumofFaces
'                        IsIn = IsInShadow(Nodep, theWorld.GetObject(q).getFace(w).getNode(1), theWorld.GetObject(q).getFace(w).getNode(2), theWorld.GetObject(q).getFace(w).getNode(3), tempL)
'                        If IsIn Then Exit For
'                    Next
'                    If IsIn Then Exit For
'                Next
'                If IsIn = False Then
                    Set MirrorP = VectorMinus(TimesVector(2, TimesVector(DotProduct(Np, tempL), Np)), tempL)
                    mirrorcosangle = VectorCosAngle(MirrorP, theWorld.ActiveCamera.CameraPoint)
            
                    Ip = Ip + theWorld.GetLight(j).Ambient  'AMBIENT LIGHT
                    If lightcosangle > 0 Then                           'DIFFUSE LIGHT
                        Ip = Ip + theWorld.GetLight(j).Intensity * Nodes(1).Parent.Kd * lightcosangle
                        If mirrorcosangle > 0 Then                          'SPECULAR LIGHT
                            Ip = Ip + theWorld.GetLight(j).Intensity * Nodes(1).Parent.Ks * mirrorcosangle ^ Nodes(1).Parent.n
                        End If
                    End If
'                Else
'                    Ip = Ip + theWorld.GetLight(j).Ambient  'AMBIENT LIGHT
'                End If
            Next
            Ip = Ip + Nodes(1).Parent.Lightness
            If Ip < 0 Then Ip = 0
            If Ip > 1 Then Ip = 1
            hls2rgb Nodes(1).Parent.hue, Ip, Nodes(1).Parent.Saturation, cR, cG, cB
'            cR = theColor.r * Ip
'            cG = theColor.g * Ip
'            cB = theColor.b * Ip
            incol = RGB(cR * 255, cG * 255, cB * 255)
            PSetAPI theWorld.RasterDevice, Xs, Ys, incol
        Next
    Next
Next
End Sub

Public Sub doFlat(RasterDevice As PictureBox, pts() As Mpoint, theColor As Long)
Dim upvertex As Integer             'to PANW shmeio tou trigwnou
Dim middlevertex As Integer         'to MESAIO shmeio tou trigwnou
Dim lowvertex As Integer            'to KATW shmeio tou trigwnou
Dim Minimun 'proswrnh metablhth sugkrishs
Dim Maximum 'proswrnh metablhth sugkrishs
Dim a(1 To 3) As Double 'to a ths eutheias
Dim b(1 To 3) As Double 'to b ths eutheias
Dim Ys As Integer   'syntetagmenh Y ths orizontias scanline eutheias
Dim Xs As Integer   'syntetagmenh X ths orizontias scanline eutheias
Dim left As Integer     'h ARISTERH pleura pou temnei h scanline
Dim right As Integer    'h DEKSIA pleura pou temnei h scanline
Dim Xa As Integer   'syntetagmenh X ths scanline me thn ARISTERH pleura
Dim Xb As Integer   'syntetagmenh X ths scanline me thn DEKSIA pleura

Minimun = 200000
Maximum = -200000
For i = 1 To 3
    If pts(i).y <= Minimun Then                 'bres thn panw
        Minimun = pts(i).y                      'kai katw korufh
        lowvertex = i
    End If
    If pts(i).y > Maximum Then
        Maximum = pts(i).y
        upvertex = i
    End If
Next
If upvertex = 1 And lowvertex = 2 Then          'twra bres kai thn
    middlevertex = 3                            'mesaia korufh
ElseIf upvertex = 1 And lowvertex = 3 Then
    middlevertex = 2
ElseIf upvertex = 2 And lowvertex = 1 Then
    middlevertex = 3
ElseIf upvertex = 2 And lowvertex = 3 Then
    middlevertex = 1
ElseIf upvertex = 3 And lowvertex = 1 Then
    middlevertex = 2
ElseIf upvertex = 3 And lowvertex = 2 Then
    middlevertex = 1
Else
    middlevertex = 1
End If

For i = 1 To 3                                  'upologise ta a,b
    k1 = i
    k2 = i + 1
    If k2 = 4 Then k2 = 1
    If pts(k1).x - pts(k2).x = 0 Then
        'katakorufh pleura ths morfhs x=b kai shmeia (b,y)
        a(i) = 1000000          'APEIRO
        b(i) = pts(k1).x        'x=b
    Else
        a(i) = (pts(k1).y - pts(k2).y) / (pts(k1).x - pts(k2).x)
        b(i) = pts(k1).y - a(i) * pts(k1).x
    End If
Next

'orise thn aristerh kai deksia pleura
If upvertex = 1 Then
    left = 3
    right = 1
ElseIf upvertex = 2 Then
    left = 1
    right = 2
ElseIf upvertex = 3 Then
    left = 2
    right = 3
End If

Y1 = pts(upvertex).y
Y2 = pts(OtherVertice(left, upvertex)).y
Y3 = pts(OtherVertice(right, upvertex)).y
For Ys = pts(upvertex).y To pts(middlevertex).y Step -1 'gia ta Y apo thn PANW
                                                        'ws th MESAIA korufh
    If a(left) = 0 Then
        'orizontia pleura
        Xa = Min(pts(middlevertex).x, pts(upvertex).x)
    ElseIf a(left) = 1000000 Then
        'katakorufh pleura
        Xa = b(left)
    Else
        Xa = (Ys - b(left)) / a(left)
    End If
    If a(right) = 0 Then
        'orizontia pleura
        Xb = Max(pts(middlevertex).x, pts(upvertex).x)
    ElseIf a(right) = 1000000 Then
        'katakorufh pleura
        Xb = b(right)
    Else
        Xb = (Ys - b(right)) / a(right)
    End If
    For Xs = Xa To Xb       'grapse ta X ths scanline
        PSetAPI RasterDevice, Xs, Ys, theColor
    Next
Next

'orise thn aristerh kai deksia pleura
If lowvertex = 1 Then
    left = 1
    right = 3
ElseIf lowvertex = 2 Then
    left = 2
    right = 1
ElseIf lowvertex = 3 Then
    left = 3
    right = 2
End If

Y1 = pts(lowvertex).y
Y2 = pts(OtherVertice(left, lowvertex)).y
Y3 = pts(OtherVertice(right, lowvertex)).y
For Ys = pts(middlevertex).y To pts(lowvertex).y Step -1    'gia ta Y apo thn MESAIA
                                                            'ws th KATW korufh
    If a(left) = 0 Then
        'orizontia pleura
        Xa = Min(pts(middlevertex).x, pts(lowvertex).x)
    ElseIf a(left) = 1000000 Then
        'katakorufh pleura
        Xa = b(left)
    Else
        Xa = (Ys - b(left)) / a(left)
    End If
    If a(right) = 0 Then
        'orizontia pleura
        Xb = Max(pts(middlevertex).x, pts(lowvertex).x)
    ElseIf a(right) = 1000000 Then
        'katakorufh pleura
        Xb = b(right)
    Else
        Xb = (Ys - b(right)) / a(right)
    End If
    For Xs = Xa To Xb       'grapse ta X ths scanline
        PSetAPI RasterDevice, Xs, Ys, theColor
    Next
Next
End Sub


'Synarthsh pou briskei thn PROHGOUMENH mias korufhs sto trigwno
Public Function PrevVertice(verticeNo As Integer) As Integer
PrevVertice = verticeNo - 1
If PrevVertice = 0 Then PrevVertice = 3
End Function
'Synarthsh pou briskei thn EPOMENH mias korufhs sto trigwno
Public Function NextVertice(verticeNo As Integer) As Integer
NextVertice = verticeNo + 1
If NextVertice = 4 Then NextVertice = 1
End Function
'Synarthsh pou briskei poia akmh kanoun 2 korufes
Public Function WhichEdge(V1 As Integer, V2 As Integer) As Integer
If (V1 = 1 And V2 = 2) Or (V1 = 2 And V2 = 1) Then
    WhichEdge = 1
ElseIf (V1 = 1 And V2 = 3) Or (V1 = 3 And V2 = 1) Then
    WhichEdge = 3
ElseIf (V1 = 2 And V2 = 3) Or (V1 = 3 And V2 = 2) Then
    WhichEdge = 2
End If
End Function
'Sunarthsh pou briskei thn allh korufh mias pleuras
Public Function OtherVertice(EdgeNo As Integer, V1 As Integer) As Integer
If EdgeNo = 1 And V1 = 1 Then
    OtherVertice = 2
ElseIf EdgeNo = 1 And V1 = 2 Then
    OtherVertice = 1
ElseIf EdgeNo = 2 And V1 = 2 Then
    OtherVertice = 3
ElseIf EdgeNo = 2 And V1 = 3 Then
    OtherVertice = 2
ElseIf EdgeNo = 3 And V1 = 3 Then
    OtherVertice = 1
ElseIf EdgeNo = 3 And V1 = 1 Then
    OtherVertice = 3
End If
End Function

Public Function IsInShadow(thePoint As Vector3D, FaceNode1 As Vector3D, FaceNode2 As Vector3D, FaceNode3 As Vector3D, theLightPoint As Vector3D) As Boolean
Dim cos1 As Double
Dim cos2 As Double
Dim cos3 As Double

cos1 = VectorCosAngle(VectorMinus(FaceNode1, thePoint), VectorMinus(theLightPoint, FaceNode1))
cos2 = VectorCosAngle(VectorMinus(FaceNode2, thePoint), VectorMinus(theLightPoint, FaceNode2))
cos3 = VectorCosAngle(VectorMinus(FaceNode3, thePoint), VectorMinus(theLightPoint, FaceNode3))

If cos1 > 0 And cos2 > 0 And cos3 > 0 Then
    IsInShadow = True
Else
    IsInShadow = False
End If
End Function
