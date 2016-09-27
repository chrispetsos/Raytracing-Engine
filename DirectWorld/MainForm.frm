VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   7800
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8790.001
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8790.001
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   4680
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6255
      Left            =   0
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Menu MENU_FILE 
      Caption         =   "File"
      Begin VB.Menu MENU_FILE_EXIT 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MENU_POPUP 
      Caption         =   "Objects"
      Begin VB.Menu MENU_FILE_OPEN 
         Caption         =   "Add"
      End
      Begin VB.Menu MENU_POPUP_MATERIAL 
         Caption         =   "Material Editor"
      End
      Begin VB.Menu MENU_POPUP_TRANSFORMATIONS 
         Caption         =   "Transformations"
      End
      Begin VB.Menu MENU_POPUP_DELETE 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu MENU_CAMERAS 
      Caption         =   "Cameras"
      Begin VB.Menu MENU_CAMERAS_ADD 
         Caption         =   "Add"
      End
      Begin VB.Menu MENU_CAMERAS_SELECT 
         Caption         =   "Select Camera"
         Begin VB.Menu MENU_CAMERAS_SELECT_CAMERA1 
            Caption         =   "Camera 1"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu MENU_CAMERAS_MANIPULATE 
         Caption         =   "Transform"
      End
      Begin VB.Menu MENU_CAMERAS_DELETE 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MENU_LIGHTS 
      Caption         =   "Lights"
      Begin VB.Menu MENU_LIGHTS_ADD 
         Caption         =   "Add"
      End
      Begin VB.Menu MENU_LIGHTS_SELECT 
         Caption         =   "Select Light"
         Begin VB.Menu MENU_LIGHTS_SELECT_LIGHT1 
            Caption         =   "Light 1"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu MENU_LIGHTS_MANIPULATE 
         Caption         =   "Manipulate"
      End
      Begin VB.Menu MENU_LIGHTS_DELETE 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MENU_RAYTRACING 
      Caption         =   "Raytracing"
      Begin VB.Menu MENU_RAYTRACING_OPTIONS 
         Caption         =   "Options"
      End
      Begin VB.Menu MENU_RAYTRACING_RAYTRACE 
         Caption         =   "Raytrace"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dragging As Boolean
Dim startDrag As D3DVECTOR2
Dim DragTo As D3DVECTOR2
Dim dvec As D3DVECTOR2
Dim d As Single

Private Sub Form_Load()
Picture1.width = Me.width
Picture1.height = Me.height

pi = 4 * Atn(1)
b = InitD3D(Picture1.hwnd)
If Not b Then
    MsgBox "Unable to CreateDevice (see InitD3D() source for comments)"
    End
End If
Timer1.enabled = True

CreateCamera vec3(1000, 1000, 1000), vec3(0, 0, 0)
CreateLight vec3(300, 300, -300)
End Sub

Private Sub Form_Resize()
'Picture1.width = Me.width
'Picture1.height = Me.height
'InitD3D Picture1.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
For Each Form In Forms
    Unload Form
Next
End Sub

Private Sub MENU_CAMERAS_ADD_Click()
CreateCamera vec3(1000, 1000, 1000), vec3(0, 0, 0)
Load MENU_CAMERAS_SELECT_CAMERA1(NumOfCameras)
MENU_CAMERAS_SELECT_CAMERA1(NumOfCameras).Caption = "Camera " & CStr(NumOfCameras)
For i = 1 To NumOfCameras
    MENU_CAMERAS_SELECT_CAMERA1(i).Checked = False
Next
MENU_CAMERAS_SELECT_CAMERA1(NumOfCameras).Checked = True
End Sub

Private Sub MENU_CAMERAS_DELETE_Click()
If NumOfCameras > 1 Then
    For i = curCamera To NumOfCameras - 1
        Set Cameras(i) = Cameras(i + 1)
    Next
    NumOfCameras = NumOfCameras - 1
    If NumOfCameras >= 1 Then
        ReDim Preserve Cameras(1 To NumOfCameras)
        If NumOfCameras = 1 Then
            curCamera = NumOfCameras
        Else
            curCamera = NumOfCameras
        End If
    End If
End If

End Sub

Private Sub MENU_CAMERAS_MANIPULATE_Click()
CameraForm.Show
End Sub

Private Sub MENU_CAMERAS_SELECT_CAMERA1_Click(Index As Integer)
curCamera = Index
For i = 1 To NumOfCameras
    MENU_CAMERAS_SELECT_CAMERA1(i).Checked = False
Next
MENU_CAMERAS_SELECT_CAMERA1(Index).Checked = True
End Sub

Private Sub MENU_FILE_EXIT_Click()
Form_Unload 0
End Sub

Private Sub MENU_FILE_OPEN_Click()
Dim newmesh As New CD3DMesh

CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
'    LoadXFile CommonDialog1.FileName
    newmesh.InitFromFile g_dev, CommonDialog1.FileName
    NumOfMeshes = NumOfMeshes + 1
    ReDim Preserve g_Mesh(1 To NumOfMeshes)
    newmesh.meshnum = NumOfMeshes
    Set g_Mesh(NumOfMeshes) = newmesh
    curmesh = NumOfMeshes
End If

End Sub

Private Sub MENU_LIGHTS_ADD_Click()
CreateLight vec3(300, 300, -300)
Load MENU_LIGHTS_SELECT_LIGHT1(NumOfLights)
MENU_LIGHTS_SELECT_LIGHT1(NumOfLights).Caption = "Light " & CStr(NumOfLights)
For i = 1 To NumOfLights
    MENU_LIGHTS_SELECT_LIGHT1(i).Checked = False
Next
MENU_LIGHTS_SELECT_LIGHT1(NumOfLights).Checked = True
End Sub

Private Sub MENU_LIGHTS_ENABLE_LIGHT1_Click(Index As Integer)

End Sub

Private Sub MENU_LIGHTS_MANIPULATE_Click()
LightForm.Show
End Sub

Private Sub MENU_LIGHTS_SELECT_LIGHT1_Click(Index As Integer)
curLight = Index
For i = 1 To NumOfLights
    MENU_LIGHTS_SELECT_LIGHT1(i).Checked = False
Next
MENU_LIGHTS_SELECT_LIGHT1(Index).Checked = True

End Sub

Private Sub MENU_POPUP_DELETE_Click()
If curmesh <> 0 Then
    For i = curmesh To NumOfMeshes - 1
        Set g_Mesh(i) = g_Mesh(i + 1)
    Next
    NumOfMeshes = NumOfMeshes - 1
    If NumOfMeshes >= 0 Then
        ReDim Preserve g_Mesh(1 To NumOfMeshes + 1)
        If NumOfMeshes = 1 Then
            curmesh = NumOfMeshes
        Else
            curmesh = NumOfMeshes
        End If
    End If
End If
End Sub

Private Sub MENU_POPUP_MATERIAL_Click()
MaterialForm.Show
End Sub

Private Sub MENU_POPUP_TRANSFORMATIONS_Click()
TransformForm.Show
End Sub

Private Sub MENU_RAYTRACING_OPTIONS_Click()
RaytraceOpts.Show
End Sub

Private Sub MENU_RAYTRACING_RAYTRACE_Click()    'Raytracing Procedure
Dim theColor As D3DCOLORVALUE
Dim pixel As D3DVECTOR2
'Dim rayOnPlane As New Vector3D
Dim rayOnPlane As D3DVECTOR
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim L As Integer
Dim fixX As Integer
Dim fixY As Integer
'Dim RayStart As New Vector3D
'Dim RayDirection As New Vector3D
Dim RayStart As D3DVECTOR
Dim RayDirection As D3DVECTOR
Dim mTime
Dim drawLine As Boolean

'ReDim RayTrace.D3DMeshes(1 To theWorld.NumberOfObjects)
'For i = 1 To theWorld.NumberOfObjects
'    ObjToMesh theWorld.GetObject(i), RayTrace.D3DMeshes(i)
''    theWorld.GetObject(i).SetTriMeshes
'Next
'Label32.Caption = "0"
'Label33.Caption = "0"
'NumOfRays = 0
RaytraceForm.Show
RaytraceForm.SetFocus

'Set theWorld.RasterDevice = RaytraceForm.Picture1
RaytraceForm.Picture1.Print "Gathering Scene and Object information..."
RaytraceForm.Picture1.Refresh
If NumOfMeshes = 0 Then Exit Sub

'theWorld.CalcRayMatrix
For i = 1 To NumOfMeshes
    g_Mesh(i).TransformAbsolute
    g_Mesh(i).GetMeshVertex
    g_Mesh(i).GetMeshIndex
'    g_Mesh(i).GetLocalBox
'    theWorld.GetObject(i).CalcVertexInFaces
'    theWorld.GetObject(i).CalcVertexNormals
Next
RaytraceForm.Picture1.Cls
RaytraceForm.Picture1.Print "      Raytracing...   "
RaytraceForm.Picture1.Refresh

mTime = Timer   'start time

'VectorToD3D theWorld.ActiveCamera.CameraPoint, RayStartD3D
RayStart = Cameras(curCamera).GetCameraPoint
'RAYTRACING VIEW-PORT SCAN
For i = 0 To RaytraceForm.Picture1.ScaleHeight Step 1
    drawLine = False
    pixel.Y = i
    RaytraceForm.Picture1.Line (0, i)-(10, i), vbRed
    For j = 0 To RaytraceForm.Picture1.ScaleWidth Step 1
        pixel.X = j
'        If theWorld.GetRayMatrix(pixel.X, pixel.Y) <> 0 Then
'            RayTrace.FaceHit = 0
            drawLine = True
            'rayOnPlane : to mono pou upologizetai xwris D3D
'            Set rayOnPlane = RayOnViewPlane(theWorld.ActiveCamera, pixel)
'            VectorToD3D rayOnPlane, rayOnPlaneD3D
'            D3DXVec3Subtract RayDirectionD3D, rayOnPlaneD3D, RayStartD3D
'            D3DXVec3Normalize RayDirectionD3D, RayDirectionD3D
'            If pixel.X = 247 And pixel.Y = 247 Then Debug.Assert False
            RayDirection = FirstRayDir(pixel)
            If TraceRay(RayStart, RayDirection, 1, theColor, 1) = 1 Then
                SetPixel RaytraceForm.Picture1.hdc, pixel.X, pixel.Y, RGB(theColor.r * 255, theColor.g * 255, theColor.b * 255)
            End If
        'End If
    Next
    If (Timer - mTime) < Int(Timer - mTime) + 0.1 And drawLine = True Then
        RaytraceForm.Picture1.Refresh   'Refresh aproximately every 1 sec
    End If
Next

'Label33.Caption = (Timer - mTime)   'end time

'Label32.Caption = CStr(NumOfRays)

'Set theWorld.RasterDevice = MainForm.Picture1
SavePicture RaytraceForm.Picture1.Image, App.Path & "\raytrace1.bmp"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tempmat As D3DMATERIAL8
Dim prevmat As D3DMATERIAL8

curmesh = ViewportPick(X, Y)
If Button = 1 Then
    dragging = True
    startDrag = vec2(X, Y)
ElseIf Button = 2 Then
    Me.PopupMenu MENU_POPUP
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If dragging Then
    DragTo = vec2(X, Y)
    D3DXVec2Subtract dvec, DragTo, startDrag
    d = D3DXVec2Length(dvec)
    If TransformMesh = True And curmesh <> 0 Then
        If startDrag.Y >= DragTo.Y Then
            PerformMeshTransformation d
        Else
            PerformMeshTransformation -d
        End If
    ElseIf TransformCamera = True Then
        If startDrag.Y >= DragTo.Y Then
            PerformCameraTransformation d
        Else
            PerformCameraTransformation -d
        End If
    ElseIf TransformLight = True Then
        If startDrag.Y >= DragTo.Y Then
            PerformLightTransformation d
        Else
            PerformLightTransformation -d
        End If
    End If
    startDrag = DragTo
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
dragging = False
End Sub

Private Sub Timer1_Timer()
Render
End Sub

Private Sub PerformMeshTransformation(Value As Single)
If TransformForm.Option1.Value = True Then
    g_Mesh(curmesh).TranslateX Value
ElseIf TransformForm.Option2.Value = True Then
    g_Mesh(curmesh).TranslateY Value
ElseIf TransformForm.Option3.Value = True Then
    g_Mesh(curmesh).TranslateZ Value
ElseIf TransformForm.Option4.Value = True Then
    g_Mesh(curmesh).RotateX Value / 100
ElseIf TransformForm.Option5.Value = True Then
    g_Mesh(curmesh).RotateY Value / 100
ElseIf TransformForm.Option6.Value = True Then
    g_Mesh(curmesh).RotateZ Value / 100
ElseIf TransformForm.Option7.Value = True Then
    g_Mesh(curmesh).ScaleUniform Value / 100
End If
End Sub

Private Sub PerformCameraTransformation(Value As Single)
If CameraForm.Option1.Value = True Then
    If CameraForm.Option3.Value = True Then
        Cameras(curCamera).AlterCameraX Value
    ElseIf CameraForm.Option4.Value = True Then
        Cameras(curCamera).AlterCameraY Value
    ElseIf CameraForm.Option5.Value = True Then
        Cameras(curCamera).AlterCameraZ Value
    End If
ElseIf CameraForm.Option2.Value = True Then
    If CameraForm.Option3.Value = True Then
        Cameras(curCamera).AlterFocusX Value
    ElseIf CameraForm.Option4.Value = True Then
        Cameras(curCamera).AlterFocusY Value
    ElseIf CameraForm.Option5.Value = True Then
        Cameras(curCamera).AlterFocusZ Value
    End If
End If
End Sub


Private Sub PerformLightTransformation(Value As Single)
'If Value > 0 Then Debug.Assert False
If LightForm.Option1.Value = True Then
    Lights(curLight).AlterX Value
ElseIf LightForm.Option2.Value = True Then
    Lights(curLight).AlterY Value
ElseIf LightForm.Option3.Value = True Then
    Lights(curLight).AlterZ Value
End If
End Sub

