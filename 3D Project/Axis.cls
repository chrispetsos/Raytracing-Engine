VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Axis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartOfAxis As New Vector3D
Public EndOfAxis As New Vector3D
Public Parent As New World
Public Name As String

Public Sub draw()
Dim StartOnRaster As New Vector2D
Dim EndOnRaster As New Vector2D
Dim startX As Integer
Dim startY As Integer
Dim endX As Integer
Dim endY As Integer

Set StartOnRaster = StartOfAxis.Perspective(Me.Parent.ActiveCamera)
StartOnRaster.FixInRasterDevice Parent.RasterDevice
startX = FixPointX(Parent.RasterDevice, StartOnRaster.x)
startY = FixPointY(Parent.RasterDevice, StartOnRaster.y)

Set EndOnRaster = EndOfAxis.Perspective(Me.Parent.ActiveCamera)
EndOnRaster.FixInRasterDevice Parent.RasterDevice
endX = FixPointX(Parent.RasterDevice, EndOnRaster.x)
endY = FixPointY(Parent.RasterDevice, EndOnRaster.y)

Parent.RasterDevice.Line (startX, startY)-(endX, endY), vbBlue
If endX >= Parent.RasterDevice.ScaleWidth / 2 Then
    endX = endX - Parent.RasterDevice.TextWidth(Name)
End If
If endY >= Parent.RasterDevice.ScaleHeight / 2 Then
    endY = endY - Parent.RasterDevice.TextHeight(Name)
End If
Parent.RasterDevice.CurrentX = endX
Parent.RasterDevice.CurrentY = endY
Parent.RasterDevice.Print Name
End Sub

