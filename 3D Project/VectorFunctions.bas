Attribute VB_Name = "VectorFunctions"
Public Function VectorDistance(Vector1 As Vector3D, Vector2 As Vector3D) As Double
VectorDistance = Sqr((Vector1.x - Vector2.x) ^ 2 + (Vector1.y - Vector2.y) ^ 2 + (Vector1.Z - Vector2.Z) ^ 2)
End Function

Public Function CrossProduct(Vector1 As Vector3D, Vector2 As Vector3D) As Vector3D
Dim tempvec As New Vector3D

tempvec.x = Vector1.y * Vector2.Z - Vector2.y * Vector1.Z
tempvec.y = Vector1.Z * Vector2.x - Vector2.Z * Vector1.x
tempvec.Z = Vector1.x * Vector2.y - Vector2.x * Vector1.y

Set CrossProduct = tempvec
End Function

Public Function DotProduct(Vector1 As Vector3D, Vector2 As Vector3D) As Double
Dim tempvec As New Vector3D

DotProduct = Vector1.x * Vector2.x + Vector1.y * Vector2.y + Vector1.Z * Vector2.Z
End Function

Public Function TimesVector(themultiplier As Double, theVector As Vector3D) As Vector3D
Dim tempvec As New Vector3D

tempvec.x = themultiplier * theVector.x
tempvec.y = themultiplier * theVector.y
tempvec.Z = themultiplier * theVector.Z
Set TimesVector = tempvec
End Function

Public Function VectorPlus(Vector1 As Vector3D, Vector2 As Vector3D) As Vector3D
Dim tempvec As New Vector3D

tempvec.x = Vector1.x + Vector2.x
tempvec.y = Vector1.y + Vector2.y
tempvec.Z = Vector1.Z + Vector2.Z

Set VectorPlus = tempvec
End Function

Public Function VectorMinus(Vector1 As Vector3D, Vector2 As Vector3D) As Vector3D
Dim tempvec As New Vector3D

tempvec.x = Vector1.x - Vector2.x
tempvec.y = Vector1.y - Vector2.y
tempvec.Z = Vector1.Z - Vector2.Z

Set VectorMinus = tempvec
End Function

Public Function VectorCosAngle(Vector1 As Vector3D, Vector2 As Vector3D) As Double
If ((Vector1.x = 0) And (Vector1.y = 0) And (Vector1.Z = 0)) Or ((Vector2.x = 0) And (Vector2.y = 0) And (Vector2.Z = 0)) Then
    VectorCosAngle = 0
Else
    VectorCosAngle = (Vector1.x * Vector2.x + Vector1.y * Vector2.y + Vector1.Z * Vector2.Z) / Sqr((Vector1.x ^ 2 + Vector1.y ^ 2 + Vector1.Z ^ 2) * (Vector2.x ^ 2 + Vector2.y ^ 2 + Vector2.Z ^ 2))
End If
End Function

Public Sub normalize(theVector As Vector3D)
Dim d As Double

d = Sqr(theVector.x ^ 2 + theVector.y ^ 2 + theVector.Z ^ 2)
If (d <> 0) Then
    d = 1 / d
    theVector.x = theVector.x * d
    theVector.y = theVector.y * d
    theVector.Z = theVector.Z * d
End If
End Sub

