Attribute VB_Name = "GeoTopo"
' GEOMETRY AND TOPOLOGY OF SHAPES
'---------------------------------

'NOTE:NOT EXCELLENT TOPOLOGY
'     OUT OF TIME
'     SORRY!!!

Option Explicit
Public koryfes() As Single
Public Nodes() As Single
Public FaceDef() As Single
Public NoOfNodes() As Single
Public edres() As Integer
Public koranaedra() As Integer
Public shmeio(4, 1) As Double
Public neoshmeio() As Double
Public nooffaces, noofvertices As Integer
Public koryfeskyklou, nokyklwn As Integer
Public koryfeshmikykl, nohmikykl As Integer
Dim i, j As Integer
Public angkyklou, angtorus, carrad As Single
Dim MArray() As Double
Dim nodenum As Integer
Dim facenum As Integer

Public Sub dokybos(pleyra As Integer)
nooffaces = 6
noofvertices = 8
ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
Nodes(1, 1) = -pleyra / 2
Nodes(1, 2) = -pleyra / 2
Nodes(1, 3) = -pleyra / 2
Nodes(2, 1) = pleyra / 2
Nodes(2, 2) = -pleyra / 2
Nodes(2, 3) = -pleyra / 2
Nodes(3, 1) = pleyra / 2
Nodes(3, 2) = pleyra / 2
Nodes(3, 3) = -pleyra / 2
Nodes(4, 1) = -pleyra / 2
Nodes(4, 2) = pleyra / 2
Nodes(4, 3) = -pleyra / 2
Nodes(5, 1) = -pleyra / 2
Nodes(5, 2) = -pleyra / 2
Nodes(5, 3) = pleyra / 2
Nodes(6, 1) = pleyra / 2
Nodes(6, 2) = -pleyra / 2
Nodes(6, 3) = pleyra / 2
Nodes(7, 1) = pleyra / 2
Nodes(7, 2) = pleyra / 2
Nodes(7, 3) = pleyra / 2
Nodes(8, 1) = -pleyra / 2
Nodes(8, 2) = pleyra / 2
Nodes(8, 3) = pleyra / 2
FaceDef(1, 1) = 1
FaceDef(1, 2) = 4
FaceDef(1, 3) = 3
FaceDef(1, 4) = 2
FaceDef(2, 1) = 5
FaceDef(2, 2) = 6
FaceDef(2, 3) = 7
FaceDef(2, 4) = 8
FaceDef(3, 1) = 1
FaceDef(3, 2) = 5
FaceDef(3, 3) = 8
FaceDef(3, 4) = 4
FaceDef(4, 1) = 2
FaceDef(4, 2) = 3
FaceDef(4, 3) = 7
FaceDef(4, 4) = 6
FaceDef(5, 1) = 1
FaceDef(5, 2) = 2
FaceDef(5, 3) = 6
FaceDef(5, 4) = 5
FaceDef(6, 1) = 3
FaceDef(6, 2) = 4
FaceDef(6, 3) = 8
FaceDef(6, 4) = 7
NoOfNodes(1) = 4
NoOfNodes(2) = 4
NoOfNodes(3) = 4
NoOfNodes(4) = 4
NoOfNodes(5) = 4
NoOfNodes(6) = 4
End Sub
Public Sub dopyramida(bash As Integer, ypsos As Integer)
nooffaces = 5
noofvertices = 5
ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
Nodes(1, 1) = -bash / 2
Nodes(1, 2) = -ypsos / 2
Nodes(1, 3) = -bash / 2
Nodes(2, 1) = bash / 2
Nodes(2, 2) = -ypsos / 2
Nodes(2, 3) = -bash / 2
Nodes(3, 1) = bash / 2
Nodes(3, 2) = -ypsos / 2
Nodes(3, 3) = bash / 2
Nodes(4, 1) = -bash / 2
Nodes(4, 2) = -ypsos / 2
Nodes(4, 3) = bash / 2
Nodes(5, 1) = 0
Nodes(5, 2) = ypsos / 2
Nodes(5, 3) = 0
FaceDef(1, 1) = 1
FaceDef(1, 2) = 4
FaceDef(1, 3) = 3
FaceDef(1, 4) = 2
FaceDef(2, 1) = 1
FaceDef(2, 2) = 2
FaceDef(2, 3) = 5
FaceDef(3, 1) = 2
FaceDef(3, 2) = 3
FaceDef(3, 3) = 5
FaceDef(4, 1) = 3
FaceDef(4, 2) = 4
FaceDef(4, 3) = 5
FaceDef(5, 1) = 4
FaceDef(5, 2) = 1
FaceDef(5, 3) = 5
NoOfNodes(1) = 4
NoOfNodes(2) = 3
NoOfNodes(3) = 3
NoOfNodes(4) = 3
NoOfNodes(5) = 3
End Sub

Public Sub doparalil(pleyra1 As Integer, pleyra2 As Integer)
nooffaces = 6
noofvertices = 8
ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
Nodes(1, 1) = -pleyra1 / 2
Nodes(1, 2) = -pleyra1 / 2
Nodes(1, 3) = -pleyra2 / 2
Nodes(2, 1) = pleyra1 / 2
Nodes(2, 2) = -pleyra1 / 2
Nodes(2, 3) = -pleyra2 / 2
Nodes(3, 1) = pleyra1 / 2
Nodes(3, 2) = pleyra1 / 2
Nodes(3, 3) = -pleyra2 / 2
Nodes(4, 1) = -pleyra1 / 2
Nodes(4, 2) = pleyra1 / 2
Nodes(4, 3) = -pleyra2 / 2
Nodes(5, 1) = -pleyra1 / 2
Nodes(5, 2) = -pleyra1 / 2
Nodes(5, 3) = pleyra2 / 2
Nodes(6, 1) = pleyra1 / 2
Nodes(6, 2) = -pleyra1 / 2
Nodes(6, 3) = pleyra2 / 2
Nodes(7, 1) = pleyra1 / 2
Nodes(7, 2) = pleyra1 / 2
Nodes(7, 3) = pleyra2 / 2
Nodes(8, 1) = -pleyra1 / 2
Nodes(8, 2) = pleyra1 / 2
Nodes(8, 3) = pleyra2 / 2
FaceDef(1, 1) = 1
FaceDef(1, 2) = 4
FaceDef(1, 3) = 3
FaceDef(1, 4) = 2
FaceDef(2, 1) = 5
FaceDef(2, 2) = 6
FaceDef(2, 3) = 7
FaceDef(2, 4) = 8
FaceDef(3, 1) = 1
FaceDef(3, 2) = 5
FaceDef(3, 3) = 8
FaceDef(3, 4) = 4
FaceDef(4, 1) = 2
FaceDef(4, 2) = 3
FaceDef(4, 3) = 7
FaceDef(4, 4) = 6
FaceDef(5, 1) = 1
FaceDef(5, 2) = 2
FaceDef(5, 3) = 6
FaceDef(5, 4) = 5
FaceDef(6, 1) = 3
FaceDef(6, 2) = 4
FaceDef(6, 3) = 8
FaceDef(6, 4) = 7
NoOfNodes(1) = 4
NoOfNodes(2) = 4
NoOfNodes(3) = 4
NoOfNodes(4) = 4
NoOfNodes(5) = 4
NoOfNodes(6) = 4
End Sub
Public Sub dokylin(aktina As Integer, pleyra As Integer, noofedres As Integer)
Dim i, k As Integer
Dim curanglex As Single
Dim curangley As Single
Dim pi As Single
pi = 4 * Atn(1)
nooffaces = noofedres
noofvertices = 2 * (nooffaces - 2)
ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
For i = 1 To noofedres - 2
    curanglex = 90 - ((360 / (noofedres - 2)) * (i - 1))
    curangley = (360 / (noofedres - 2)) * (i - 1)
    curanglex = curanglex * pi / 180
    curangley = curangley * pi / 180
    Nodes(i, 1) = aktina * Cos(curanglex)
    Nodes(i, 2) = aktina * Cos(curangley)
    Nodes(i, 3) = -pleyra / 2
    Nodes(i + noofvertices / 2, 1) = aktina * Cos(curanglex)
    Nodes(i + noofvertices / 2, 2) = aktina * Cos(curangley)
    Nodes(i + noofvertices / 2, 3) = pleyra / 2
Next
For i = 1 To noofedres - 3
    FaceDef(i, 1) = i
    FaceDef(i, 2) = i + 1
    FaceDef(i, 3) = i + 1 + noofvertices / 2
    FaceDef(i, 4) = i + noofvertices / 2
Next
FaceDef(noofedres - 2, 1) = noofvertices / 2
FaceDef(noofedres - 2, 2) = 1
FaceDef(noofedres - 2, 3) = noofvertices / 2 + 1
FaceDef(noofedres - 2, 4) = noofvertices
For i = 1 To noofvertices / 2
    FaceDef(noofedres - 1, i) = noofvertices / 2 - i + 1
Next
For i = 1 To noofvertices / 2
    FaceDef(noofedres, i) = noofvertices / 2 + i
Next
For i = 1 To noofedres
    If i <= noofedres - 2 Then NoOfNodes(i) = 4 Else NoOfNodes(i) = noofvertices / 2
Next
End Sub
Public Sub rotatey(x As Single, y As Single, Z As Single, ByVal angle As Double)
shmeio(1, 1) = x
shmeio(2, 1) = y
shmeio(3, 1) = Z
shmeio(4, 1) = 1
PrepareRotationY MArray, angle
MultiplyArray MArray, shmeio, neoshmeio
End Sub

Public Sub translate(ByVal Tx As Single, ByVal Ty As Single, ByVal Tz As Single)
'shmeio(1, 1) = x
'shmeio(2, 1) = y
'shmeio(3, 1) = Z
shmeio(4, 1) = 1
'PrepareTranslation MArray, Tx, Ty, Tz
'MultiplyArray shmeio, MArray, neoshmeio
End Sub

Public Sub dotorus(paxos As Integer, aktina As Integer, noofedres As Integer)
Dim pi As Single

nodenum = 0
facenum = 0
pi = 4 * Atn(1)
nooffaces = noofedres
koryfeskyklou = nooffaces / 20
nokyklwn = nooffaces / koryfeskyklou
noofvertices = koryfeskyklou * nokyklwn
angkyklou = 360 / koryfeskyklou
angtorus = 360 / nokyklwn
angkyklou = angkyklou * pi / 180
angtorus = angtorus * pi / 180

ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
For i = 1 To koryfeskyklou
    Nodes(i, 1) = aktina + paxos / 2 + (paxos / 2) * Sin(angkyklou * (i - 1))
    Nodes(i, 2) = (paxos / 2) * Cos(angkyklou * (i - 1))
    Nodes(i, 3) = 0
Next
For j = 1 To nokyklwn - 1
    For i = 1 To koryfeskyklou
        rotatey Nodes(i + (j - 1) * koryfeskyklou, 1), Nodes(i + (j - 1) * koryfeskyklou, 2), Nodes(i + (j - 1) * koryfeskyklou, 3), angtorus
        Nodes(i + koryfeskyklou * j, 1) = neoshmeio(1, 1)
        Nodes(i + koryfeskyklou * j, 2) = neoshmeio(2, 1)
        Nodes(i + koryfeskyklou * j, 3) = neoshmeio(3, 1)
    Next
Next
For j = 1 To nokyklwn
    For i = 1 To koryfeskyklou - 1
        FaceDef(i + (koryfeskyklou * (j - 1)), 1) = i + (koryfeskyklou * (j - 1))
        FaceDef(i + (koryfeskyklou * (j - 1)), 2) = i + (koryfeskyklou * (j - 1)) + 1
        FaceDef(i + (koryfeskyklou * (j - 1)), 3) = i + (koryfeskyklou * (j - 1)) + 1 + koryfeskyklou
        FaceDef(i + (koryfeskyklou * (j - 1)), 4) = i + (koryfeskyklou * (j - 1)) + 1 + koryfeskyklou - 1
    Next
    FaceDef(koryfeskyklou * j, 1) = koryfeskyklou * j
    FaceDef(koryfeskyklou * j, 2) = koryfeskyklou * (j - 1) + 1
    FaceDef(koryfeskyklou * j, 3) = koryfeskyklou * (j - 1) + 1 + koryfeskyklou
    FaceDef(koryfeskyklou * j, 4) = koryfeskyklou * j + koryfeskyklou
Next
For i = nooffaces - koryfeskyklou + 1 To nooffaces
    FaceDef(i, 3) = FaceDef(i, 3) - nooffaces
    FaceDef(i, 4) = FaceDef(i, 4) - nooffaces
Next
For i = 1 To nooffaces
    NoOfNodes(i) = 4
Next

Open App.Path + "\MyTorus.txt" For Output As #1   ' Open file for output.
Print #1, noofvertices, nooffaces
For i = 1 To noofvertices
    nodenum = nodenum + 1
    Print #1, nodenum, Int(Nodes(i, 1)), Int(Nodes(i, 2)), Int(Nodes(i, 3))
Next
For i = 1 To nooffaces
    facenum = facenum + 1
    Print #1, facenum, "4", FaceDef(i, 1), FaceDef(i, 2), FaceDef(i, 3), FaceDef(i, 4)
Next
Close #1
End Sub

Public Sub dosphere(aktina As Integer, noofedres As Integer)
Dim pi, angsphere As Single, anghmikykl As Single
pi = 4 * Atn(1)
nodenum = 0
facenum = 0
nooffaces = noofedres
koryfeshmikykl = nooffaces / 20
nohmikykl = nooffaces / koryfeshmikykl
noofvertices = koryfeshmikykl * nohmikykl
anghmikykl = 180 / (koryfeshmikykl - 1)
angsphere = 360 / nohmikykl
anghmikykl = anghmikykl * pi / 180
angsphere = angsphere * pi / 180
ReDim Nodes(nooffaces, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
For i = 1 To koryfeshmikykl + 1
    Nodes(i, 1) = aktina * Sin(anghmikykl * (i - 1))
    Nodes(i, 2) = aktina * Cos(anghmikykl * (i - 1))
    Nodes(i, 3) = 0
Next
For j = 1 To nohmikykl - 1
    For i = 1 To koryfeshmikykl
        rotatey Nodes(i + (j - 1) * koryfeshmikykl, 1), Nodes(i + (j - 1) * koryfeshmikykl, 2), Nodes(i + (j - 1) * koryfeshmikykl, 3), angsphere
        Nodes(i + koryfeshmikykl * j, 1) = neoshmeio(1, 1)
        Nodes(i + koryfeshmikykl * j, 2) = neoshmeio(2, 1)
        Nodes(i + koryfeshmikykl * j, 3) = neoshmeio(3, 1)
    Next
Next
For j = 1 To nohmikykl - 1
    For i = 1 To koryfeshmikykl - 1
        FaceDef(i + (koryfeshmikykl * (j - 1)), 1) = i + (koryfeshmikykl * (j - 1))
        FaceDef(i + (koryfeshmikykl * (j - 1)), 2) = i + (koryfeshmikykl * (j - 1)) + 1
        FaceDef(i + (koryfeshmikykl * (j - 1)), 3) = i + (koryfeshmikykl * (j - 1)) + 1 + koryfeshmikykl
        FaceDef(i + (koryfeshmikykl * (j - 1)), 4) = i + (koryfeshmikykl * (j - 1)) + 1 + koryfeshmikykl - 1
    Next
    FaceDef(koryfeshmikykl * j, 1) = koryfeshmikykl
    FaceDef(koryfeshmikykl * j, 2) = koryfeshmikykl
    FaceDef(koryfeshmikykl * j, 3) = koryfeshmikykl
    FaceDef(koryfeshmikykl * j, 4) = koryfeshmikykl
Next
For i = nooffaces - koryfeshmikykl + 1 To nooffaces - 1
    FaceDef(i, 1) = i
    FaceDef(i, 2) = i + 1
    FaceDef(i, 3) = i + 1 - nooffaces + koryfeshmikykl
    FaceDef(i, 4) = i - nooffaces + koryfeshmikykl
Next
FaceDef(nooffaces, 1) = koryfeshmikykl
FaceDef(nooffaces, 2) = koryfeshmikykl
FaceDef(nooffaces, 3) = koryfeshmikykl
FaceDef(nooffaces, 4) = koryfeshmikykl
For i = 1 To noofvertices
    NoOfNodes(i) = 4
Next

Open App.Path + "\MySphere.txt" For Output As #1   ' Open file for output.
Print #1, noofvertices, nooffaces
For i = 1 To noofvertices
    nodenum = nodenum + 1
    Print #1, nodenum, Int(Nodes(i, 1)), Int(Nodes(i, 2)), Int(Nodes(i, 3))
Next
For i = 1 To nooffaces
    facenum = facenum + 1
    Print #1, facenum, "4", FaceDef(i, 1), FaceDef(i, 2), FaceDef(i, 3), FaceDef(i, 4)
Next
Close #1
End Sub

Public Sub dokonos(ypsos As Integer, aktina As Integer, noofedres As Integer)
Dim pi As Single
pi = 4 * Atn(1)
nooffaces = noofedres
noofvertices = nooffaces
koryfeskyklou = noofvertices - 1
angkyklou = 360 / koryfeskyklou
angkyklou = angkyklou * pi / 180
ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
For i = 1 To koryfeskyklou
    Nodes(i, 1) = aktina * Sin(angkyklou * (i - 1))
    Nodes(i, 2) = aktina * Cos(angkyklou * (i - 1))
    Nodes(i, 3) = -ypsos / 2
Next
Nodes(koryfeskyklou + 1, 1) = 0
Nodes(koryfeskyklou + 1, 2) = 0
Nodes(koryfeskyklou + 1, 3) = ypsos / 2
For i = 1 To koryfeskyklou
    FaceDef(1, i) = koryfeskyklou - i + 1
Next
For i = 2 To nooffaces - 1
    FaceDef(i, 1) = i - 1
    FaceDef(i, 2) = i
    FaceDef(i, 3) = koryfeskyklou + 1
Next
FaceDef(koryfeskyklou + 1, 1) = koryfeskyklou
FaceDef(koryfeskyklou + 1, 2) = 1
FaceDef(koryfeskyklou + 1, 3) = koryfeskyklou + 1
NoOfNodes(1) = koryfeskyklou
For i = 2 To nooffaces
    NoOfNodes(i) = 3
Next
End Sub
Public Sub doglass()
Dim pi As Single
Dim i As Integer
Dim ry(20, 2) As Integer
nodenum = 0
facenum = 0
ry(1, 1) = 100
ry(1, 2) = -100
ry(2, 1) = 90
ry(2, 2) = -90
ry(3, 1) = 70
ry(3, 2) = -70
ry(4, 1) = 60
ry(4, 2) = -60
ry(5, 1) = 40
ry(5, 2) = -50
ry(6, 1) = 30
ry(6, 2) = -40
ry(7, 1) = 20
ry(7, 2) = -30
ry(8, 1) = 17
ry(8, 2) = -20
ry(9, 1) = 13
ry(9, 2) = -10
ry(10, 1) = 10
ry(10, 2) = 0
ry(11, 1) = 40
ry(11, 2) = 10
ry(12, 1) = 70
ry(12, 2) = 20
ry(13, 1) = 100
ry(13, 2) = 30
ry(14, 1) = 110
ry(14, 2) = 40
ry(15, 1) = 120
ry(15, 2) = 50
ry(16, 1) = 130
ry(16, 2) = 70
ry(17, 1) = 135
ry(17, 2) = 90
ry(18, 1) = 140
ry(18, 2) = 110
ry(19, 1) = 145
ry(19, 2) = 130
ry(20, 1) = 150
ry(20, 2) = 150
pi = 4 * Atn(1)
koryfeskyklou = 20
nokyklwn = 20
nooffaces = koryfeskyklou * nokyklwn - koryfeskyklou '+ 2
noofvertices = koryfeskyklou * nokyklwn
angkyklou = 360 / koryfeskyklou
angkyklou = angkyklou * pi / 180
ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
For j = 1 To nokyklwn
    For i = 1 To koryfeskyklou
        Nodes(i + (j - 1) * koryfeskyklou, 1) = ry(j, 1) * Sin(angkyklou * (i - 1))
        Nodes(i + (j - 1) * koryfeskyklou, 2) = ry(j, 2)
        Nodes(i + (j - 1) * koryfeskyklou, 3) = ry(j, 1) * Cos(angkyklou * (i - 1))
    Next i
Next j
For j = 1 To nokyklwn - 1
    For i = 1 To koryfeskyklou - 1
        FaceDef(i + (koryfeskyklou * (j - 1)), 1) = i + (koryfeskyklou * (j - 1))
        FaceDef(i + (koryfeskyklou * (j - 1)), 2) = i + (koryfeskyklou * (j - 1)) + 1
        FaceDef(i + (koryfeskyklou * (j - 1)), 3) = i + (koryfeskyklou * (j - 1)) + 1 + koryfeskyklou
        FaceDef(i + (koryfeskyklou * (j - 1)), 4) = i + (koryfeskyklou * (j - 1)) + 1 + koryfeskyklou - 1
    Next
    FaceDef(koryfeskyklou * j, 1) = koryfeskyklou * j
    FaceDef(koryfeskyklou * j, 2) = koryfeskyklou * (j - 1) + 1
    FaceDef(koryfeskyklou * j, 3) = koryfeskyklou * (j - 1) + 1 + koryfeskyklou
    FaceDef(koryfeskyklou * j, 4) = koryfeskyklou * j + koryfeskyklou
Next
For i = 1 To nooffaces '- 2
    NoOfNodes(i) = 4
Next
'NoOfNodes(nooffaces - 1) = koryfeskyklou
'NoOfNodes(nooffaces) = koryfeskyklou

Open App.Path + "\glass.txt" For Output As #1   ' Open file for output.
Print #1, noofvertices, nooffaces
For i = 1 To noofvertices
    nodenum = nodenum + 1
    Print #1, nodenum, Int(Nodes(i, 1)), Int(Nodes(i, 2)), Int(Nodes(i, 3))
Next
For i = 1 To nooffaces ' - 2
    facenum = facenum + 1
    Print #1, facenum, NoOfNodes(i), FaceDef(i, 1), FaceDef(i, 2), FaceDef(i, 3), FaceDef(i, 4)
Next
'facenum = facenum + 1
'Print #1, facenum, NoOfNodes(i), FaceDef(i, 1), FaceDef(i, 2), FaceDef(i, 3), FaceDef(i, 4)
Close #1
End Sub
Public Sub dobottle()
Dim pi As Single
Dim i As Integer
Dim ry(17, 2) As Single
nodenum = 0
facenum = 0
ry(1, 1) = 70
ry(1, 2) = -200
ry(2, 1) = 70
ry(2, 2) = -150
ry(3, 1) = 70
ry(3, 2) = -75
ry(4, 1) = 70
ry(4, 2) = -50
ry(5, 1) = 70
ry(5, 2) = 0
ry(6, 1) = 67
ry(6, 2) = 30
ry(7, 1) = 65
ry(7, 2) = 50
ry(8, 1) = 55
ry(8, 2) = 65
ry(9, 1) = 40
ry(9, 2) = 75
ry(10, 1) = 40
ry(10, 2) = 85
ry(11, 1) = 30
ry(11, 2) = 100
ry(12, 1) = 20
ry(12, 2) = 115
ry(13, 1) = 20
ry(13, 2) = 130
ry(14, 1) = 25
ry(14, 2) = 135
ry(15, 1) = 27
ry(15, 2) = 140
ry(16, 1) = 27
ry(16, 2) = 145
ry(17, 1) = 25
ry(17, 2) = 150
pi = 4 * Atn(1)
koryfeskyklou = 20
nokyklwn = 17
nooffaces = koryfeskyklou * nokyklwn - koryfeskyklou '+ 2
noofvertices = koryfeskyklou * nokyklwn
angkyklou = 360 / koryfeskyklou
angkyklou = angkyklou * pi / 180
ReDim Nodes(noofvertices, 3)
ReDim FaceDef(nooffaces, 1000)
ReDim NoOfNodes(nooffaces)
For i = 1 To nooffaces '- 2
    NoOfNodes(i) = 4
Next
'NoOfNodes(nooffaces - 1) = koryfeskyklou
'NoOfNodes(nooffaces) = koryfeskyklou

For j = 1 To nokyklwn
    For i = 1 To koryfeskyklou
        Nodes(i + (j - 1) * koryfeskyklou, 1) = ry(j, 1) * Sin(angkyklou * (i - 1))
        Nodes(i + (j - 1) * koryfeskyklou, 2) = ry(j, 2)
        Nodes(i + (j - 1) * koryfeskyklou, 3) = ry(j, 1) * Cos(angkyklou * (i - 1))
    Next i
Next j
For j = 1 To nokyklwn - 1
    For i = 1 To koryfeskyklou - 1
        FaceDef(i + (koryfeskyklou * (j - 1)), 1) = i + (koryfeskyklou * (j - 1))
        FaceDef(i + (koryfeskyklou * (j - 1)), 2) = i + (koryfeskyklou * (j - 1)) + 1
        FaceDef(i + (koryfeskyklou * (j - 1)), 3) = i + (koryfeskyklou * (j - 1)) + 1 + koryfeskyklou
        FaceDef(i + (koryfeskyklou * (j - 1)), 4) = i + (koryfeskyklou * (j - 1)) + 1 + koryfeskyklou - 1
    Next
    FaceDef(koryfeskyklou * j, 1) = koryfeskyklou * j
    FaceDef(koryfeskyklou * j, 2) = koryfeskyklou * (j - 1) + 1
    FaceDef(koryfeskyklou * j, 3) = koryfeskyklou * (j - 1) + 1 + koryfeskyklou
    FaceDef(koryfeskyklou * j, 4) = koryfeskyklou * j + koryfeskyklou
Next

Open App.Path + "\Bottle.txt" For Output As #1   ' Open file for output.
Print #1, noofvertices, nooffaces
For i = 1 To noofvertices
    nodenum = nodenum + 1
    Print #1, nodenum, Int(Nodes(i, 1)), Int(Nodes(i, 2)), Int(Nodes(i, 3))
Next
For i = 1 To nooffaces ' - 2
    facenum = facenum + 1
    Print #1, facenum, NoOfNodes(i), FaceDef(i, 1), FaceDef(i, 2), FaceDef(i, 3), FaceDef(i, 4)
Next
Close #1
End Sub

Public Sub rearrays()
ReDim PerNodes(noofvertices, 3) As Single
ReDim painterZ(nooffaces, 3) As Single
ReDim visicol(nooffaces, 2) As Single
ReDim Painter(nooffaces) As Integer
ReDim FaceDepth(nooffaces) As Single
End Sub

Public Sub doTerrain()
nooffaces = 25
noofvertices = 36
ReDim Nodes(1 To noofvertices, 1 To 3)
ReDim FaceDef(1 To nooffaces, 1 To 4)
ReDim NoOfNodes(nooffaces)
For i = 1 To 6
    For j = 1 To 6
        Nodes(6 * (i - 1) + j, 1) = 10 * (i - 1)
        Nodes(6 * (i - 1) + j, 2) = 10 * (j - 1)
        Nodes(6 * (i - 1) + j, 3) = 0
    Next
Next
For i = 1 To 5
    For j = 1 To 5
        FaceDef(5 * (i - 1) + j, 1) = 5 * (i - 1) + j
        FaceDef(5 * (i - 1) + j, 2) = 5 * i + j
        FaceDef(5 * (i - 1) + j, 3) = 5 * i + j + 1
        FaceDef(5 * (i - 1) + j, 4) = 5 * (i - 1) + j + 1
    Next
Next
nodenum = 0
facenum = 0
Open App.Path + "\myTerrain.txt" For Output As #1   ' Open file for output.
Print #1, noofvertices, nooffaces
For i = 1 To noofvertices
    nodenum = nodenum + 1
    Print #1, nodenum, Int(Nodes(i, 1)), Int(Nodes(i, 2)), Int(Nodes(i, 3))
Next
For i = 1 To nooffaces ' - 2
    facenum = facenum + 1
    Print #1, facenum, 4, FaceDef(i, 1), FaceDef(i, 2), FaceDef(i, 3), FaceDef(i, 4)
Next
Close #1
End Sub

