Attribute VB_Name = "ArrayFuncs"
Public SinThetaY As Double, CosThetaY As Double, SinThetaX As Double, CosThetaX As Double, pi As Double

Public Sub MultiplyArray(MArray1() As Double, MArray2() As Double, MArray3() As Double)
    Dim i As Integer, j As Integer, k As Integer
    ReDim MArray3(1 To UBound(MArray1, 1), 1 To UBound(MArray2, 2))
    For j = 1 To UBound(MArray3, 1)
        For k = 1 To UBound(MArray3, 2)
            For i = 1 To UBound(MArray1, 2)
                MArray3(j, k) = MArray3(j, k) + (MArray1(j, i) * MArray2(i, k))
            Next
        Next
    Next
End Sub

Public Sub PrepareArray(MArray() As Double, Rows As Integer, Columns As Integer, ParamArray arguments())
    Dim i As Integer, j As Integer, k As Integer
    i = 0
    ReDim MArray(1 To Rows, 1 To Columns)
    For j = 1 To Rows
        For k = 1 To Columns
            MArray(j, k) = arguments(i)
            i = i + 1
        Next
    Next
End Sub

Public Sub AssignArray(MArray1() As Double, MArray2() As Double)
    Dim j As Integer, k As Integer
    ReDim MArray1(1 To UBound(MArray2, 1), 1 To UBound(MArray2, 2))
    For j = 1 To UBound(MArray2, 1)
        For k = 1 To UBound(MArray2, 2)
            MArray1(j, k) = MArray2(j, k)
        Next
    Next
End Sub

Public Sub PrepareScaling(MArray() As Double, Sx As Double, Sy As Double, Sz As Double)
    PrepareArray MArray, 4, 4, Sx, 0, 0, 0, 0, Sy, 0, 0, 0, 0, Sz, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareTranslation(MArray() As Double, Tx As Double, Ty As Double, Tz As Double)
    PrepareArray MArray, 4, 4, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, Tx, Ty, Tz, 1
End Sub

Public Sub PrepareRotationZ(MArray() As Double, Theta As Double)
    PrepareArray MArray, 4, 4, Cos(Theta), Sin(Theta), 0, 0, -Sin(Theta), Cos(Theta), 0, 0, 0, 0, 1, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareRotationY(MArray() As Double, Theta As Double)
    PrepareArray MArray, 4, 4, Cos(Theta), 0, -Sin(Theta), 0, 0, 1, 0, 0, Sin(Theta), 0, Cos(Theta), 0, 0, 0, 0, 1
End Sub

Public Sub PrepareRotationX(MArray() As Double, Theta As Double)
    PrepareArray MArray, 4, 4, 1, 0, 0, 0, 0, Cos(Theta), Sin(Theta), 0, 0, -Sin(Theta), Cos(Theta), 0, 0, 0, 0, 1
End Sub

Public Sub PrepareReflectionXY(MArray() As Double)
    PrepareArray MArray, 4, 4, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareReflectionXZ(MArray() As Double)
    PrepareArray MArray, 4, 4, 1, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareReflectionYZ(MArray() As Double)
    PrepareArray MArray, 4, 4, -1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareReflection0(MArray() As Double)
    PrepareArray MArray, 4, 4, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareShearing(MArray() As Double, SHxy As Double, SHxz As Double, SHyx As Double, SHyz As Double, SHzx As Double, SHzy As Double)
    PrepareArray MArray, 4, 4, 1, SHxy, SHxz, 0, SHyx, 1, SHyz, 0, SHzx, SHzy, 1, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareOrthographicXY(MArray() As Double)
    PrepareArray MArray, 4, 4, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1
End Sub

Public Sub PrepareIsometricXY(MArray() As Double)
    PrepareArray MArray, 4, 4, CosThetaY, SinThetaY * SinThetaX, 0, 0, 0, CosThetaX, 0, 0, SinThetaY, -SinThetaX * CosThetaX, 0, 0, 0, 0, 0, 1
End Sub


