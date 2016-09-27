VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private CameraPoint As D3DVECTOR
Private FocusPoint As D3DVECTOR
Public CameraNum As Integer

Public Sub Create(theCameraPoint As D3DVECTOR, theFocusPoint As D3DVECTOR)
CameraPoint = theCameraPoint
FocusPoint = theFocusPoint
End Sub

Public Sub AlterCameraZ(Value As Single)
CameraPoint.z = CameraPoint.z + Value
End Sub

Public Sub AlterCameraY(Value As Single)
CameraPoint.Y = CameraPoint.Y + Value
End Sub

Public Sub AlterCameraX(Value As Single)
CameraPoint.X = CameraPoint.X + Value
End Sub

Public Sub AlterFocusZ(Value As Single)
FocusPoint.z = FocusPoint.z + Value
End Sub

Public Sub AlterFocusY(Value As Single)
FocusPoint.Y = FocusPoint.Y + Value
End Sub

Public Sub AlterFocusX(Value As Single)
FocusPoint.X = FocusPoint.X + Value
End Sub

Public Property Get GetCameraPoint() As D3DVECTOR
GetCameraPoint = CameraPoint
End Property

Public Property Get GetFocusPoint() As D3DVECTOR
GetFocusPoint = FocusPoint
End Property

