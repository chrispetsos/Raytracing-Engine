Attribute VB_Name = "General"
Public ZeroVector As New Vector3D
Public pi As Double

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Sub StartCode()
pi = 4 * Atn(1)
ZeroVector.X = 0
ZeroVector.Y = 0
ZeroVector.Z = 0
End Sub

Public Function DegsToRads(degrees As Double) As Double
    DegsToRads = (degrees * pi) / 180
End Function

Public Function CreateObjectFromFile(theFileName As String, theColor As FaceColor) As Object3D
Dim tempobj As New Object3D
Dim NoOfNodes(10000) As Integer
Dim theNodes() As Integer
Dim noofvertices As Integer
Dim nooffaces As Integer
Dim j As Integer
Dim X As Double
Dim Y As Double
Dim Z As Double

On Error GoTo ErrorHandler
    If theFileName <> "" Then
    Open theFileName For Input As #1
    Input #1, noofvertices, nooffaces
    tempobj.Create noofvertices, nooffaces, 100, 0.1, 0.5, 0.5, 0.5, 10
    For i = 1 To noofvertices
        Input #1, j, X, Y, Z
        tempobj.AddNode j, X, Y, Z
    Next
    For i = 1 To nooffaces
        Input #1, j
        Input #1, NoOfNodes(j)
        ReDim Preserve theNodes(1 To NoOfNodes(j))
        For k = 1 To NoOfNodes(j)
            Input #1, theNodes(k)
        Next
        tempobj.AddFace j, NoOfNodes(j), theNodes
    Next
    Close #1
    Set CreateObjectFromFile = tempobj
    End If
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 32755
            Exit Function
        Case Else
            Err.Raise Err.Number
    End Select
End Function

Public Sub PSetAPI(RasterDevice As PictureBox, X As Integer, Y As Integer, Color As Long)
SetPixel RasterDevice.hdc, FixPointX(RasterDevice, X), FixPointY(RasterDevice, Y), Color
End Sub

Public Function FixPointX(RasterDevice As PictureBox, ByVal X As Long) As Long
FixPointX = X + RasterDevice.ScaleWidth / 2
End Function

Public Function FixPointY(RasterDevice As PictureBox, ByVal Y As Long) As Long
FixPointY = RasterDevice.ScaleHeight / 2 - Y
End Function

Public Sub Tesselate(theObject As Object3D)
Dim tempobj As New Object3D
Dim facenum As Integer

facenum = 0
For i = 1 To theObject.NumofFaces
    If theObject.getFace(i).NumofNodes > 3 Then
        For j = 1 To theObject.getFace(i).NumofNodes - 2
            facenum = facenum + 1
        Next
    Else
        facenum = facenum + 1
    End If
Next
tempobj.Create theObject.NumofNodes, facenum, theObject.hue, theObject.Lightness, theObject.Saturation, theObject.Kd, theObject.Ks, theObject.n
facenum = 0
For i = 1 To theObject.NumofFaces
    If theObject.getFace(i).NumofNodes > 3 Then
        For j = 1 To theObject.getFace(i).NumofNodes - 2
            facenum = facenum + 1
            tempobj.AddFace facenum, 3, theObject.getFace(i).getNodeNo(1), theObject.getFace(i).getNodeNo(j + 1), theObject.getFace(i).getNodeNo(j + 2)
        Next
    Else
        facenum = facenum + 1
        tempobj.AddFace facenum, 3, theObject.getFace(i).getNodeNo(1), theObject.getFace(i).getNodeNo(2), theObject.getFace(i).getNodeNo(3)
    End If
Next
For i = 1 To theObject.NumofNodes
    tempobj.AddNode i, theObject.getNode(i).X, theObject.getNode(i).Y, theObject.getNode(i).Z
Next
Set theObject = tempobj
End Sub

Public Function IsZeroVector(theVector As Vector3D) As Boolean
If theVector.X = 0 And theVector.Y = 0 And theVector.Z = 0 Then
    IsZeroVector = True
Else
    IsZeroVector = False
End If
End Function

Public Function PointInRectangle(ByVal PointX As Integer, ByVal PointY As Integer, ByVal RectMinX As Integer, ByVal RectMaxX As Integer, ByVal RectMinY As Integer, ByVal RectMaxY As Integer) As Boolean
PointInRectangle = False
If PointX > RectMinX And PointX < RectMaxX And PointY > RectMinY And PointY < RectMaxY Then
    PointInRectangle = True
End If

End Function

Public Sub delay(Milliseconds As Long)
Dim myTime As Date
myTime = Now
Do While Now - myTime < Milliseconds / 1E+15
Loop
End Sub

Public Function FixFiveDigit(theNum As String) As String
If Len(theNum) = 0 Then
    FixFiveDigit = "00000"
ElseIf Len(theNum) = 1 Then
    FixFiveDigit = "0000" & theNum
ElseIf Len(theNum) = 2 Then
    FixFiveDigit = "000" & theNum
ElseIf Len(theNum) = 3 Then
    FixFiveDigit = "00" & theNum
ElseIf Len(theNum) = 4 Then
    FixFiveDigit = "0" & theNum
ElseIf Len(theNum) = 5 Then
    FixFiveDigit = "" & theNum
End If
End Function
