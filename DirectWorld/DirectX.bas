Attribute VB_Name = "DirectX"
Public g_dx As New DirectX8
Public g_d3d As Direct3D8              'Used to create the D3DDevice
Public g_d3dx As New D3DX8
Public g_dev As Direct3DDevice8  'Our rendering device
Dim mode As D3DDISPLAYMODE
Dim d3dpp As D3DPRESENT_PARAMETERS
'Dim g_MeshMaterials() As D3DMATERIAL8   ' Mesh Material data
'Dim g_MeshTextures() As Direct3DTexture8 ' Mesh Textures
'Dim g_NumMaterials As Long

Public pi As Single

Public g_Mesh() As New CD3DMesh                  ' Our Meshes
Public curmesh As Integer
Public NumOfMeshes As Integer
Public TransformMesh As Boolean

Public Cameras() As New Camera
Public curCamera As Integer
Public NumOfCameras As Integer
Public TransformCamera As Boolean

Public Lights() As New Light
Public curLight As Integer
Public NumOfLights As Integer
Public TransformLight As Boolean

' A structure for our custom vertex type
' representing a point on the screen
Public Type CUSTOMVERTEX
    X As Single         'x in screen space
    Y As Single         'y in screen space
    z  As Single        'normalized z
    normal As D3DVECTOR       'vertex normal
'    v As Single       'vertex color
End Type

Public Type CUSTOMINDEX
    v1 As Integer
    v2 As Integer
    v3 As Integer
End Type

Public Type CUSTOMLINE
    v1 As Integer
    v2 As Integer
End Type

Public Type CUSTOMQUADFACE
    v1 As Integer
    v2 As Integer
    v3 As Integer
    v4 As Integer
End Type

Private Type D3DXINTERSECTINFO
    FaceIndex As Long
    u As Single
    v As Single
    Dist As Single
End Type

Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL)

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


'-----------------------------------------------------------------------------
' Name: InitD3D()
' Desc: Initializes Direct3D
'-----------------------------------------------------------------------------
Function InitD3D(hwnd As Long) As Boolean
    Dim worldMat As D3DMATERIAL8
    
    On Local Error Resume Next
    
    ' Create the D3D object
    Set g_d3d = g_dx.Direct3DCreate()
    If g_d3d Is Nothing Then Exit Function
    
     ' Get The current Display Mode format
    Dim mode As D3DDISPLAYMODE
    g_d3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
         
    ' Set up the structure used to create the D3DDevice. Since we are now
    ' using more complex geometry, we will create a device with a zbuffer.
    ' the D3DFMT_D16 indicates we want a 16 bit z buffer but
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = 1
    d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
    d3dpp.BackBufferFormat = mode.Format
    d3dpp.BackBufferCount = 1
    d3dpp.EnableAutoDepthStencil = 1
    d3dpp.AutoDepthStencilFormat = D3DFMT_D16

    ' Create the D3DDevice
    ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
    ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
    
    Set g_dev = g_d3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, _
                                    D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If g_dev Is Nothing Then Exit Function
    
    ' Device state would normally be set here
    ' Turn off culling, so we see the front and back of the triangle
    g_dev.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    ' Turn on the zbuffer
    g_dev.SetRenderState D3DRS_ZENABLE, 1
    
    ' Turn on lighting
'    g_dev.SetRenderState D3DRS_LIGHTING, 1
    
    ' Turn on full ambient light to white
'    g_dev.SetRenderState D3DRS_AMBIENT, &HFFFFFFFF
    
'    g_dev.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
'    g_dev.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
'    g_dev.SetRenderState D3DRS_FILLMODE, D3DFILL_POINT

'     g_dev.SetRenderState D3DRS_AMBIENTMATERIALSOURCE, D3DMCS_COLOR1
'    worldMat.Ambient.r = 1
'    worldMat.Ambient.g = 1
'    worldMat.Ambient.b = 1
'    worldMat.Ambient.a = 1
'    g_dev.SetMaterial worldMat
    g_dev.SetRenderState D3DRS_AMBIENT, RGB(255, 255, 255)
    g_dev.SetRenderState D3DRS_SPECULARENABLE, 1
'    g_dev.SetRenderState D3DRS_WRAP0, D3DWRAPCOORD_0 Or D3DWRAPCOORD_1
    RayTraceDepth = 2
    WeightTreshold = 0.1
    InitD3D = True
End Function

'-----------------------------------------------------------------------------
' Name: Render()
' Desc: Draws the scene
'-----------------------------------------------------------------------------
Sub Render()

    Dim i As Long
    
    If g_dev Is Nothing Then Exit Sub

    ' Clear the backbuffer to a blue color (ARGB = 000000ff)
    ' Clear the z buffer to 1
    g_dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    
     
    ' Setup the world, view, and projection matrices
    SetupMatrices
    SetupLights
'    g_dev.SetVertexShader D3DFVF_CUSTOMVERTEX
    
    ' Begin the scene
    g_dev.BeginScene
    
    ' Meshes are divided into subsets, one for each material.
    ' Render them in a loop

'    For i = 0 To g_NumMaterials - 1
    
        ' Set the material and texture for this subset
'        g_dev.SetMaterial g_MeshMaterials(i)
'        g_dev.SetTexture 0, g_MeshTextures(i)
        'draw the lights
        For j = 1 To NumOfLights
            Lights(j).draw
        Next
        
        ' Draw the meshes subset
        For j = 1 To NumOfMeshes
            g_Mesh(j).Render g_dev
        Next
'    Next
            
    ' End the scene
    g_dev.EndScene
    
     
    ' Present the backbuffer contents to the front buffer (screen)
    g_dev.Present ByVal 0, ByVal 0, 0, ByVal 0
    
End Sub

Public Sub CreateCamera(CameraPoint As D3DVECTOR, FocusPoint As D3DVECTOR)
Dim newCamera As New Camera

newCamera.Create CameraPoint, FocusPoint
NumOfCameras = NumOfCameras + 1
newCamera.CameraNum = NumOfCameras
ReDim Preserve Cameras(1 To NumOfCameras)
Set Cameras(NumOfCameras) = newCamera
curCamera = NumOfCameras


End Sub

Public Sub CreateLight(LightPoint As D3DVECTOR)
Dim newLight As New Light

newLight.Create LightPoint
NumOfLights = NumOfLights + 1
newLight.LightNum = NumOfLights
ReDim Preserve Lights(1 To NumOfLights)
Set Lights(NumOfLights) = newLight
curLight = NumOfLights
g_dev.SetLight NumOfLights - 1, newLight.GetD3DLight  'let d3d know about the light
g_dev.LightEnable NumOfLights - 1, 1                  'turn it on
g_dev.SetRenderState D3DRS_LIGHTING, 1    'make sure lighting is enabled
End Sub

'-----------------------------------------------------------------------------
' Name: SetupMatrices()
' Desc: Sets up the world, view, and projection transform matrices.
'-----------------------------------------------------------------------------
Sub SetupMatrices()

    
    ' The transform Matrix is used to position and orient the objects
    ' you are drawing
    ' For our world matrix, we will just rotate the object about the y axis.
'    Dim matWorld As D3DMATRIX
'    D3DXMatrixRotationAxis matWorld, vec3(0, 1, 0), 0.5
'    g_dev.SetTransform D3DTS_WORLD, matWorld
'    g_dev.SetTransform

    ' The view matrix defines the position and orientation of the camera
    ' Set up our view matrix. A view matrix can be defined given an eye point,
    ' a point to lookat, and a direction for which way is up. Here, we set the
    ' eye five units back along the z-axis and up three units, look at the
    ' origin, and define "up" to be in the y-direction.
    
    
    Dim matView As D3DMATRIX
'    D3DXMatrixLookAtLH matView, vec3(500#, 400#, 300#), _
                                 vec3(0#, 0#, 0#), _
                                 vec3(0#, 1#, 0#)
    D3DXMatrixLookAtLH matView, Cameras(curCamera).GetCameraPoint, _
                                Cameras(curCamera).GetFocusPoint, _
                                vec3(0#, 1#, 0#)
                                 
    g_dev.SetTransform D3DTS_VIEW, matView

    ' The projection matrix describes the camera's lenses
    ' For the projection matrix, we set up a perspective transform (which
    ' transforms geometry from 3D view space to 2D viewport space, with
    ' a perspective divide making objects smaller in the distance). To build
    ' a perpsective transform, we need the field of view (1/4 pi is common),
    ' the aspect ratio, and the near and far clipping planes (which define at
    ' what distances geometry should be no longer be rendered).
    Dim matProj As D3DMATRIX
    D3DXMatrixPerspectiveFovLH matProj, pi / 4, 1, 1, 100000
    g_dev.SetTransform D3DTS_PROJECTION, matProj

End Sub



'-----------------------------------------------------------------------------
' Name: SetupLights()
' Desc: Sets up the lights and materials for the scene.
'-----------------------------------------------------------------------------
Sub SetupLights()
     
'    Dim col As D3DCOLORVALUE
    
    
    ' Set up a material. The material here just has the diffuse and ambient
    ' colors set to yellow. Note that only one material can be used at a time.
'    Dim mtrl As D3DMATERIAL8
'    With col:    .r = 1: .g = 1: .b = 0: .a = 1:   End With
'    mtrl.diffuse = col
'    mtrl.Ambient = col
'    g_dev.SetMaterial mtrl
    
    ' Set up a white, directional light, with an oscillating direction.
    ' Note that many lights may be active at a time (but each one slows down
    ' the rendering of our scene). However, here we are just using one. Also,
    ' we need to set the D3DRS_LIGHTING renderstate to enable lighting
    
'    Dim Light As D3DLIGHT8
'    Light.Type = D3DLIGHT_POINT
'    Light.diffuse.r = 0.1
'    Light.diffuse.g = 0.1
'    Light.diffuse.b = 0.1
'    Light.specular.r = 1
'    light.Direction.X = -1#
'    light.Direction.Y = -1#
'    light.Direction.z = -1#
'    Light.Position = vec3(1000, 1000, 1000)
'    Light.Range = 100000#
    
'    g_dev.SetLight 0, Light                   'let d3d know about the light
'    g_dev.LightEnable 0, 1                    'turn it on
'    g_dev.SetRenderState D3DRS_LIGHTING, 1    'make sure lighting is enabled

    ' Finally, turn on some ambient light.
    ' Ambient light is light that scatters and lights all objects evenly
'    g_dev.SetRenderState D3DRS_AMBIENT, 0
'    g_dev.SetRenderState D3DRS_SPECULARENABLE, 1
'    g_dev.SetRenderState D3DRS_AMBIENTMATERIALSOURCE, D3DMCS_COLOR2
End Sub

'-----------------------------------------------------------------------------
' Name: vec3()
' Desc: helper function
'-----------------------------------------------------------------------------
Function vec2(X As Single, Y As Single) As D3DVECTOR2
    vec2.X = X
    vec2.Y = Y
End Function
Function vec3(X As Single, Y As Single, z As Single) As D3DVECTOR
    vec3.X = X
    vec3.Y = Y
    vec3.z = z
End Function

'-----------------------------------------------------------------------------
' Name: ViewportPick
' Params:
'    frame      parent of frame heirarchy to pick from
'    x          x screen coordinate in pixels
'    y          y screen coordinate in pixels
'
' Note: After call GetCount to see if any objets where hit
'-----------------------------------------------------------------------------
Public Function ViewportPick(X As Single, Y As Single) As Integer
    Dim viewport As D3DVIEWPORT8
    Dim world As D3DMATRIX
    Dim proj As D3DMATRIX
    Dim view As D3DMATRIX
    
    'NOTE the following functions will fail on PURE HAL devices
    'use ViewportPickEx if working with pureHal devices
    
    g_dev.GetViewport viewport
'    world = g_identityMatrix
    D3DXMatrixIdentity world
    g_dev.GetTransform D3DTS_VIEW, view
    g_dev.GetTransform D3DTS_PROJECTION, proj
    
    ViewportPick = ViewportPickEx(viewport, proj, view, world, X, Y)
    
End Function


'-----------------------------------------------------------------------------
' Name: ViewportPickEx
' Desc: Aux function for ViewportPick
'-----------------------------------------------------------------------------
Public Function ViewportPickEx(viewport As D3DVIEWPORT8, proj As D3DMATRIX, view As D3DMATRIX, world As D3DMATRIX, X As Single, Y As Single) As Integer
    
'    If frame.Enabled = False Then Exit Function
    
    Dim vIn As D3DVECTOR, vNear As D3DVECTOR, vFar As D3DVECTOR, vDir As D3DVECTOR
    Dim bHit As Integer, i As Long
    Dim minraydist As Double
    Dim retHit As Long
    Dim retfaceindextriFaceid As Long
    Dim u As Single
    Dim v As Single
    Dim retdist As Single
    Dim countHits As Long
    Dim buf As D3DXBuffer
'    If frame Is Nothing Then Exit Function
    minraydist = 2000000
                        
    Dim currentMatrix As D3DMATRIX
    Dim NewWorldMatrix As D3DMATRIX
    
'    currentMatrix = frame.GetMatrix
    currentMatrix = world
    'Setup our basis matrix for this frame
    D3DXMatrixMultiply NewWorldMatrix, currentMatrix, world
    
    vIn.X = X:    vIn.Y = Y
    
    'Compute point on Near Clip plane at cursor
    vIn.z = 0
    D3DXVec3Unproject vNear, vIn, viewport, proj, view, NewWorldMatrix
    
    'compute point on far clip plane at cursor
    vIn.z = 1
    D3DXVec3Unproject vFar, vIn, viewport, proj, view, NewWorldMatrix

    'Comput direction vector
    D3DXVec3Subtract vDir, vFar, vNear
    D3DXVec3Normalize vDir, vDir
    
    
'    Dim item As D3D_PICK_RECORD
    
    
    'Check all child meshes
    'Even if we got a hit we continue as the next mesh may be closer
    Dim childMesh As CD3DMesh
    For i = 1 To NumOfMeshes
        
        Set childMesh = g_Mesh(i)
        childMesh.TransformAbsolute
        If Not childMesh Is Nothing Then
            Set buf = g_d3dx.Intersect(childMesh.TransformedMesh, vNear, vDir, retHit, retfaceindextriFaceid, u, v, retdist, countHits)
'            Set buf = g_d3dx.Intersect(childMesh.mesh, vNear, vDir, retHit, retfaceindextriFaceid, u, v, retdist, countHits)
        End If
        
'        If item.hit <> 0 Then
'            InternalAddItem frame, childMesh, item
'            item.hit = 0
'        End If
        If Not (buf Is Nothing) Then
            If retdist < minraydist Then
                minraydist = retdist
                bHit = i
            End If
        End If
    Next
    
    'check pick for all child frame
'    Dim childFrame As CD3DFrame
'    For i = 0 To frame.GetChildFrameCount() - 1
'        Set childFrame = frame.GetChildFrame(i)
'        bHit = bHit Or _
'                ViewportPickEx(childFrame, viewport, proj, view, NewWorldMatrix, X, Y)
'    Next
'
    ViewportPickEx = bHit

End Function


'-----------------------------------------------------------------------------
' Name: LONGtoD3DCOLORVALUE
'-----------------------------------------------------------------------------
Public Function LONGtoD3DCOLORVALUE(color As Long) As D3DCOLORVALUE
    Dim a As Long, r As Long, g As Long, b As Long
        
    If color < 0 Then
        a = ((color And (&H7F000000)) / (2 ^ 24)) Or &H80&
    Else
        a = color / (2 ^ 24)
    End If
    b = (color And &HFF0000) / (2 ^ 16)
    g = (color And &HFF00&) / (2 ^ 8)
    r = (color And &HFF&)
    
    LONGtoD3DCOLORVALUE.a = a / 255
    LONGtoD3DCOLORVALUE.r = r / 255
    LONGtoD3DCOLORVALUE.g = g / 255
    LONGtoD3DCOLORVALUE.b = b / 255
        
End Function

Public Function FirstRayDir(thePixel As D3DVECTOR2) As D3DVECTOR
    Dim viewport As D3DVIEWPORT8
    Dim world As D3DMATRIX
    Dim proj As D3DMATRIX
    Dim view As D3DMATRIX
    Dim vIn As D3DVECTOR, vNear As D3DVECTOR, vFar As D3DVECTOR, vDir As D3DVECTOR
    Dim bHit As Integer, i As Long
    Dim minraydist As Double
    Dim retHit As Long
    Dim retfaceindextriFaceid As Long
    Dim u As Single
    Dim v As Single
    Dim retdist As Single
    Dim countHits As Long
    Dim buf As D3DXBuffer
        
    Dim currentMatrix As D3DMATRIX
    Dim NewWorldMatrix As D3DMATRIX
    
    'NOTE the following functions will fail on PURE HAL devices
    'use ViewportPickEx if working with pureHal devices
    
    g_dev.GetViewport viewport
'    world = g_identityMatrix
    D3DXMatrixIdentity world
    g_dev.GetTransform D3DTS_VIEW, view
    g_dev.GetTransform D3DTS_PROJECTION, proj
    
'    If frame Is Nothing Then Exit Function
'    minraydist = 2000000
                           
'    currentMatrix = frame.GetMatrix
    currentMatrix = world
    'Setup our basis matrix for this frame
    D3DXMatrixMultiply NewWorldMatrix, currentMatrix, world
    
    vIn.X = thePixel.X:    vIn.Y = thePixel.Y
    
    'Compute point on Near Clip plane at cursor
    vIn.z = 0
    D3DXVec3Unproject vNear, vIn, viewport, proj, view, NewWorldMatrix
    
    'compute point on far clip plane at cursor
    vIn.z = 1
    D3DXVec3Unproject vFar, vIn, viewport, proj, view, NewWorldMatrix

    'Comput direction vector
    D3DXVec3Subtract vDir, vFar, vNear
    D3DXVec3Normalize vDir, vDir
    
    FirstRayDir = vDir
End Function


Public Function IntersectMesh(theMesh As CD3DMesh, RayStart As D3DVECTOR, RayDir As D3DVECTOR, retdist As Single, InterP As D3DVECTOR, InterPnormal As D3DVECTOR, retFaceIndex As Long) As Boolean
Dim retHit As Long
'Dim retFaceIndex As Long
Dim u As Single
Dim v As Single
Dim countHits As Long
Dim buf As D3DXBuffer
Dim temp3d As D3DVECTOR
Dim temp3d1 As D3DVECTOR
Dim temp3d2 As D3DVECTOR
Dim temp3d3 As D3DVECTOR
Dim theData As D3DXINTERSECTINFO
Dim minraydist As Double


minraydist = 2000000
Set buf = g_d3dx.Intersect(theMesh.TransformedMesh, RayStart, RayDir, retHit, retFaceIndex, u, v, retdist, countHits)

'inter = V1 + U(V2-V1) + V(V3-V1)
'Any point in the plane V1V2V3 can be represented by the barycentric coordinate
'(U,V). The parameter U controls how much V2 gets weighted into the result
'and the parameter V controls how much V3 gets weighted into the result.
'Lastly, 1-U-V controls how much V1 gets weighted into the result.
i = 0
If Not (buf Is Nothing) Then        'if intersections exist
    Do
        i = i + 1       'find closest to ray but not too close in order
                        ' to avoid self reflection
        g_d3dx.BufferGetData buf, i - 1, Len(theData), 1, theData
        If theData.Dist < minraydist And theData.Dist > 0.001 Then
            D3DXVec3Scale temp3d, RayDir, theData.Dist
            D3DXVec3Add InterP, RayStart, temp3d
'            D3DXVec3Scale temp3d1, theObj.getFace(theData.FaceIndex + 1).getNode(1).GetNormalD3D, 1 - theData.u - theData.v
'            D3DXVec3Scale temp3d2, theObj.getFace(theData.FaceIndex + 1).getNode(2).GetNormalD3D, theData.u
'            D3DXVec3Scale temp3d3, theObj.getFace(theData.FaceIndex + 1).getNode(3).GetNormalD3D, theData.v
            D3DXVec3Scale temp3d1, theMesh.GetVertexNormal(theMesh.GetIndex(theData.FaceIndex, 1)), 1 - theData.u - theData.v
            D3DXVec3Scale temp3d2, theMesh.GetVertexNormal(theMesh.GetIndex(theData.FaceIndex, 2)), theData.u
            D3DXVec3Scale temp3d3, theMesh.GetVertexNormal(theMesh.GetIndex(theData.FaceIndex, 3)), theData.v
            D3DXVec3Add temp3d, temp3d1, temp3d2
            D3DXVec3Add InterPnormal, temp3d, temp3d3
            minraydist = theData.Dist
            retdist = minraydist
            retFaceIndex = theData.FaceIndex '+ 1
            IntersectMesh = True
        End If
    Loop Until i = countHits
'    If retdist < 0.001 Then
'        IntersectMesh = False
'    Else
'        IntersectMesh = True
'    End If
End If
End Function

