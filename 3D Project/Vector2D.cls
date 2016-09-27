VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public X As Double
Public Y As Double

Public Sub FixInRasterDevice(RasterDevice As PictureBox)
Dim a As Double

If (Abs(X) > RasterDevice.ScaleWidth / 2 Or Abs(Y) > RasterDevice.ScaleHeight / 2) Then
    If X <> 0 Then
        a = Y / X
        If Abs(X) > RasterDevice.ScaleWidth / 2 Then
            If X > 0 Then
                X = RasterDevice.ScaleWidth / 2
                Y = a * RasterDevice.ScaleWidth / 2
            Else
                X = -RasterDevice.ScaleWidth / 2
                Y = -a * RasterDevice.ScaleWidth / 2
            End If
        End If
        If Abs(Y) > RasterDevice.ScaleHeight / 2 Then
            If Y > 0 Then
                X = (RasterDevice.ScaleHeight / 2) / a
                Y = RasterDevice.ScaleHeight / 2
            Else
                X = -(RasterDevice.ScaleHeight / 2) / a
                Y = -RasterDevice.ScaleHeight / 2
            End If
        End If
    Else
        If Y > 0 Then
            Y = RasterDevice.ScaleHeight / 2
        Else
            Y = -RasterDevice.ScaleHeight / 2
        End If
    End If
End If
End Sub
