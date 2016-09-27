VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CD3DMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Copyright (C) 1999-2001 Microsoft Corporation.  All Rights Reserved.
'
'  File:       D3DMesh.cls
'  Content:    D3D VB Framework Mesh
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Private Type BoneComboIdList
    List(4) As Long
End Type


' Base Objects
Public mesh As D3DXMesh                         'if not skinned, regular mesh object
Public TransformedMesh As D3DXMesh                         'if not skinned, regular mesh object
Public skinmesh As D3DXSkinMesh                 'if skinned - skinned mesh object
Public bUseMaterials As Boolean                 'Use materials in object
Public bUseMaterialOverride As Boolean          'Use only override material
Public ObjectName As String                     'Name of object
Public bSkinned As Boolean                      'Inidicates if the object is a skin

' Culling objects
Dim m_SphereCenter As D3DVECTOR                 'center of bounding sphere
Public SphereRadius As Single                   'radius of bounding sphere
Public bHasSphere As Boolean                    'Inidcatges if bounding sphere is calculated

' Material and Adjacency information
Dim m_MaterialOverride As D3DMATERIAL8          'Override material to use if bUseMaterialOverride set
Dim m_NumMaterials As Long                      'Number of materials in object
Dim m_Materials() As D3DMATERIAL8               'Material List
Dim m_Textures() As Direct3DTexture8            'Texture List
Dim m_TextureNames() As String                  'Texture List Names
Dim m_AdjBuffer As D3DXBuffer                   'Adjacency buffer for the mesh


' Bone and skinning informatin
Dim m_BoneNames As D3DXBuffer                   'Names of Frame objects that are bones
Dim m_BoneMatrices As D3DXBuffer                'Matrix object per bone
Dim m_BoneComboTable As D3DXBuffer              'Groupings of bone material and triangles
'Dim m_BoneFrames() As CD3DFrame                 'Frame objects that are bones
Dim m_maxFaceInfl As Long                       'Number of matrices that will influence a vertex
Dim m_worldMatHandles(4) As Long                'handle to multiple world matrices
Dim m_BoneOffsetMat() As D3DMATRIX              'Bone offset matrices from D3DXBuffers
Dim m_NumBoneCombos As Long                     'Size of bonecombo table
Dim m_BoneCombos() As D3DXBONECOMBINATION       'Combo table as returned from D3DX
Dim m_BoneComboIds() As BoneComboIdList         'BoneId portion of combotable
Dim m_BoneCount As Long                         'Number of bones
Dim m_bonesAttached As Boolean                  'Indicates if bones have been attached to the mesh

'  Transformation Objects
Dim TranslationVector As D3DVECTOR
Dim RotationVector As D3DVECTOR
Dim ScalingVector As D3DVECTOR
Public meshnum As Long
Private BoundBoxVertex() As D3DVECTOR
Private BoundBoxIndex() As CUSTOMLINE
Private BoundBox As D3DXMesh
Private BoundBoxVB As Direct3DVertexBuffer8
Private BoundBoxIB As Direct3DIndexBuffer8

Private minExt As D3DVECTOR
Private maxExt As D3DVECTOR
Private transminExt As D3DVECTOR
Private transmaxExt As D3DVECTOR

Private Vertex() As CUSTOMVERTEX
Private Index() As CUSTOMINDEX

Public Kd As Double
Public Ks As Double
Public n As Double
Public h As Double
Public Kl As Double
Public Ktg As Double
Public Krg As Double
'-----------------------------------------------------------------------------
' Init
'-----------------------------------------------------------------------------
Sub Init()
    bUseMaterials = True
End Sub

'-----------------------------------------------------------------------------
' SphereCenter()
' returns Sphere Center since D3DVECTOR can not be public variable return value
'-----------------------------------------------------------------------------
Property Get SphereCenter() As D3DVECTOR
    SphereCenter = m_SphereCenter
End Property

'-----------------------------------------------------------------------------
' Name: InitFromFile
' Desc: Tries first to load mesh in current directory or using explicit path
'       If that fails loads mesh from SDK media path
'-----------------------------------------------------------------------------

Public Function InitFromFile(dev As Direct3DDevice8, Name As String) As Boolean
    Dim mtrlBuffer As D3DXBuffer
    Dim strPath As String
    Destroy
    
'    On Local Error Resume Next
    Set m_AdjBuffer = Nothing
    bSkinned = False
    
    Set mesh = g_d3dx.LoadMeshFromX(Name, D3DXMESH_MANAGED, dev, m_AdjBuffer, mtrlBuffer, m_NumMaterials)
'    Set TransformedMesh = mesh
    Set TransformedMesh = mesh.CloneMeshFVF(D3DXMESH_VB_DYNAMIC, D3DFVF_XYZ, g_dev)
    If Err.Number <> 0 Then
        Err.Clear
        On Local Error GoTo errOut
'        strPath = g_mediaPath + Name
'        Set mesh = g_d3dx.LoadMeshFromX(strPath, D3DXMESH_MANAGED, dev, m_AdjBuffer, mtrlBuffer, m_NumMaterials)
    
    End If
    
    Call InitMaterials(g_dev, mtrlBuffer)
    
    TranslationVector = vec3(0, 0, 0)
    RotationVector = vec3(0, 0, 0)
    ScalingVector = vec3(1, 1, 1)
    
    ReDim BoundBoxIndex(11)
    BoundBoxIndex(0).v1 = 0
    BoundBoxIndex(0).v2 = 1
    BoundBoxIndex(1).v1 = 1
    BoundBoxIndex(1).v2 = 3
    BoundBoxIndex(2).v1 = 3
    BoundBoxIndex(2).v2 = 2
    BoundBoxIndex(3).v1 = 2
    BoundBoxIndex(3).v2 = 0
    BoundBoxIndex(4).v1 = 5
    BoundBoxIndex(4).v2 = 4
    BoundBoxIndex(5).v1 = 4
    BoundBoxIndex(5).v2 = 6
    BoundBoxIndex(6).v1 = 6
    BoundBoxIndex(6).v2 = 7
    BoundBoxIndex(7).v1 = 7
    BoundBoxIndex(7).v2 = 5
    BoundBoxIndex(8).v1 = 3
    BoundBoxIndex(8).v2 = 7
    BoundBoxIndex(9).v1 = 1
    BoundBoxIndex(9).v2 = 5
    BoundBoxIndex(10).v1 = 0
    BoundBoxIndex(10).v2 = 4
    BoundBoxIndex(11).v1 = 2
    BoundBoxIndex(11).v2 = 6
    
    Kd = 0.8
    Ks = 0
    n = 10
    h = 1
    Kl = 1
    Ktg = 0
    Krg = 0
'    Set BoundBox = g_d3dx.CreateMeshFVF(6, 8, D3DXMESH_POINTS, D3DFVF_XYZ, g_dev)
    
'    Set BoundBoxIB = g_dev.CreateIndexBuffer(Len(BoundBoxIndex(0)) * 6, D3DUSAGE_DYNAMIC, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
'    Set BoundBoxIB = BoundBox.GetIndexBuffer
'    BoundBoxIB.Lock 0, 0, 0, 0
'    D3DIndexBuffer8SetData BoundBoxIB, 0, Len(BoundBoxIndex(0)) * 6, D3DLOCK_DISCARD, BoundBoxIndex(0)
'    BoundBoxIB.Unlock
    InitFromFile = True
    Exit Function
    
errOut:
    InitFromFile = False
End Function



'-----------------------------------------------------------------------------
' Name: InitFromXOF
' Desc: Load mesh from data provided by XOF api
'       Called from D3DUtil_LoadFromFile
'-----------------------------------------------------------------------------

'Public Function InitFromXOF(dev As Direct3DDevice8, meshdata As DirectXFileData) As Boolean
'    Dim mtrlBuffer As D3DXBuffer
'    Dim bonename As String
'    Dim i As Long
'    Dim q As Long
        
        
'    Destroy
    
    
'    Set m_AdjBuffer = Nothing
'    Set m_BoneMatrices = Nothing
'    Set m_BoneNames = Nothing
'    Set mesh = Nothing
'    Set skinmesh = Nothing
'    Set m_BoneMatrices = Nothing
'    Set m_BoneComboTable = Nothing
    
'    ObjectName = meshdata.GetName()
    
''    On Local Error GoTo errOut
'    bSkinned = False
    
'    'Users can set this variable to TRUE try the skinned load path
'    If g_bLoadSkins = True Then
'        Set skinmesh = g_d3dx.LoadSkinMeshFromXof(meshdata, D3DXMESH_MANAGED, dev, m_AdjBuffer, mtrlBuffer, m_NumMaterials, m_BoneNames, m_BoneMatrices)
        
        
'        Dim pAdj As Long, AdjOut As D3DXBuffer
'        pAdj = m_AdjBuffer.GetBufferPointer
        
'        m_BoneCount = skinmesh.GetNumBones()
        
'        If m_BoneCount = 0 Then
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' a skinned mesh with no bones is just a regular mesh
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            
'            bSkinned = False
'            Set mesh = skinmesh.GetOriginalMesh()
'            'Set skinmesh = Nothing
            
'        Else
        
            '''''''''''''''''''''''''''''''''''''''''''''''
            ' code specific to x files with skinning data in them
            '''''''''''''''''''''''''''''''''''''''''''''''
            
'            bSkinned = True
            
'            Set mesh = skinmesh.ConvertToBlendedMesh(D3DXMESH_SYSTEMMEM, ByVal pAdj, ByVal 0, m_NumBoneCombos, m_BoneComboTable, ByVal 0&, Nothing)
'            Set m_AdjBuffer = Nothing
'            Set m_AdjBuffer = AdjOut
'            Set AdjOut = Nothing
            
            'retrieve number of influence (matrices) that a vertices could have
            'we support up to 4 corresponding to the 4 world matrices that can be set
'            m_maxFaceInfl = skinmesh.GetMaxFaceInfluences()
            
'            m_worldMatHandles(0) = D3DTS_WORLD
'            m_worldMatHandles(1) = D3DTS_WORLD1
'            m_worldMatHandles(2) = D3DTS_WORLD2
'            m_worldMatHandles(3) = D3DTS_WORLD3
            
'            ReDim m_BoneCombos(m_NumBoneCombos)
'            ReDim m_BoneComboIds(m_NumBoneCombos)
            
            ' fill in our private table for bone combo data
            ' this inidicates which bones (matrices) need to be blended
            ' for a given subset in the mesh
'            For q = 0 To m_NumBoneCombos - 1
'                g_d3dx.BufferGetBoneCombo m_BoneComboTable, q, m_BoneCombos(q)
'                g_d3dx.BufferGetBoneComboBoneIds m_BoneComboTable, q, m_maxFaceInfl, m_BoneComboIds(q).List(0)
'            Next
'            Set m_BoneComboTable = Nothing
                    
            ' fill in our private table for bone offset matrices
            ' these are the matrices that give the intitial displacement of mesh subsets
            ' release the d3dx buffer to save memory
'            ReDim m_BoneOffsetMat(m_BoneCount)
'            g_d3dx.BufferGetData m_BoneMatrices, 0, Len(m_BoneOffsetMat(0)), m_BoneCount, m_BoneOffsetMat(0)
'            Set m_BoneMatrices = Nothing
            
            
'        End If
'    Else
    
'         Set mesh = g_d3dx.LoadMeshFromXof(meshdata, D3DXMESH_MANAGED, dev, m_AdjBuffer, mtrlBuffer, m_NumMaterials)
         
'    End If
    
    
'    Call InitMaterials(g_dev, mtrlBuffer)
    
'    InitFromXOF = True
'    Exit Function
    
'errOut:
'    InitFromXOF = False
'End Function


'-----------------------------------------------------------------------------
' Name: AttatchBonesToMesh
' Desc: Called to attach bones to a skin.
'       The BoneNames table is used to search out bone frames
'       in the children of the given parent frame
'
'       This must be done for any skinning animation to work
'-----------------------------------------------------------------------------

'Friend Sub AttatchBonesToMesh(parent As CD3DFrame)
'      ' get links to all the frames (bones)
'        Dim i As Long
'        Dim bonename As String
'        ReDim m_BoneFrames(m_BoneCount)
'        For i = 0 To m_BoneCount - 1
'            bonename = g_d3dx.BufferGetBoneName(m_BoneNames, i)
'            Set m_BoneFrames(i) = parent.FindChildObject(bonename, 0)
'            If m_BoneFrames(i) Is Nothing Then
'                Debug.Print "unable to find " + bonename
'                Stop
'            End If
'        Next
'        m_bonesAttached = True
'        Set m_BoneNames = Nothing
'End Sub


'-----------------------------------------------------------------------------
' Name: Optimize
' Desc: Re-organize the mesh for better performance
'
'-----------------------------------------------------------------------------

Sub Optimize()
    Dim s As Long
    Dim adjBuf1() As Long
    Dim adjBuf2() As Long
    Dim facemap() As Long
    Dim newmesh As D3DXMesh
    Dim vertexMap As D3DXBuffer
    
    s = m_AdjBuffer.GetBufferSize
    ReDim adjBuf1(s / 4)
    ReDim adjBuf2(s / 4)
    
    s = mesh.GetNumFaces
    ReDim facemap(s)
    
    g_d3dx.BufferGetData m_AdjBuffer, 0, 4, s * 3, adjBuf1(0)
    
    Set newmesh = mesh.Optimize(D3DXMESHOPT_ATTRSORT Or D3DXMESHOPT_VERTEXCACHE, adjBuf1(0), adjBuf2(0), facemap(0), vertexMap)
    
    If Not newmesh Is Nothing Then
        Set mesh = Nothing
        Set mesh = newmesh
    End If

End Sub




'-----------------------------------------------------------------------------
' Name: InitMaterials
' Desc: Helper function for creating mesh materials
'       called after initialization
'-----------------------------------------------------------------------------
Private Sub InitMaterials(d3ddevice As Direct3DDevice8, mtrlBuffer As D3DXBuffer)
    Dim i As Long
    
    If m_NumMaterials <= 0 Then Exit Sub
    
    ReDim m_Materials(m_NumMaterials)
    ReDim m_Textures(m_NumMaterials)
    ReDim m_TextureNames(m_NumMaterials)

    For i = 0 To m_NumMaterials - 1
        'copy material out of material buffer into our own structure
        g_d3dx.BufferGetMaterial mtrlBuffer, i, m_Materials(i)

'        If g_bLoadNoAlpha Then m_Materials(i).diffuse.a = 1
        
'        m_Materials(i).Ambient = m_Materials(i).diffuse
        
        m_TextureNames(i) = g_d3dx.BufferGetTextureName(mtrlBuffer, i)
        
'        If g_bUseTextureLoadCallback Then
'            Set m_Textures(i) = g_TextureLoadCallback.TextureLoadCallback(m_TextureNames(i))
'        Else
'            Set m_Textures(i) = D3DUtil_CreateTextureInPool(g_dev, m_TextureNames(i), D3DFMT_UNKNOWN)
'        End If
    Next
    
    
End Sub



'-----------------------------------------------------------------------------
' Name: SetFVF
' Desc: Change the FVF of the current mesh
'----------------------------------------------------------------------------
Public Sub SetFVF(dev As Direct3DDevice8, fvf As Long)
    
    Dim tempMesh As D3DXMesh
    Dim verts() As D3DVERTEX

    If mesh Is Nothing Then Exit Sub
    
    Set tempMesh = mesh.CloneMeshFVF(D3DXMESH_MANAGED, fvf, dev)
    
        
    Set mesh = tempMesh
    
    
End Sub

'-----------------------------------------------------------------------------
' Name: GenerateNormals
' Desc: if the current mesh Flexible Vertex Format (FVF) has normals in it
'       that are not initialized. This function will fill them.
'       if no normals are present in the FVF this function will fire an
'       exception
'----------------------------------------------------------------------------
Public Sub ComputeNormals()
    Dim bm As D3DXBaseMesh
    Set bm = mesh
    g_d3dx.ComputeNormals bm
End Sub


'-----------------------------------------------------------------------------
' Name: FlipNormals
' Desc: Convenience function that flips normals for a D3DVERTEX mesh (default)
'----------------------------------------------------------------------------
Public Sub FlipNormals()
    Dim count As Long
    Dim size As Long
    Dim i As Long
    
    Dim verts() As D3DVERTEX
    
    Dim vb As Direct3DVertexBuffer8
    Set vb = mesh.GetVertexBuffer()
    
    
    
    
    size = g_d3dx.GetFVFVertexSize(mesh.GetFVF())
    count = mesh.GetNumVertices()
    
    
    
    If mesh.GetFVF() = D3DFVF_VERTEX Then
        ReDim verts(count)
        
        D3DVertexBuffer8GetData vb, 0, size * count, 0, verts(0)
        
        For i = 0 To count - 1
            verts(i).nx = -verts(i).nx
            verts(i).ny = -verts(i).ny
            verts(i).nz = -verts(i).nz
        Next
        
        D3DVertexBuffer8SetData vb, 0, size * count, 0, verts(0)
    Else
        Stop
    End If
    

End Sub


'-----------------------------------------------------------------------------
' Name: Translate
' Desc: all vertices are moved by x,y,z
'       note that object will still rotate about 0,0,0
'
'----------------------------------------------------------------------------
Public Sub TranslateAbsolute(X As Single, Y As Single, z As Single)
    Dim count As Long
    Dim size As Long
    Dim i As Long
    
    Dim verts() As D3DVERTEX
    
    Dim vb As Direct3DVertexBuffer8
    Set vb = mesh.GetVertexBuffer()
    
    
    
    size = g_d3dx.GetFVFVertexSize(mesh.GetFVF())
    count = mesh.GetNumVertices()

    
    If mesh.GetFVF() = D3DFVF_VERTEX Then
        ReDim verts(count)
        
        D3DVertexBuffer8GetData vb, 0, size * count, 0, verts(0)
        
        For i = 0 To count - 1
            verts(i).X = verts(i).X + X
            verts(i).Y = verts(i).Y + Y
            verts(i).z = verts(i).z + z
        Next
        
        D3DVertexBuffer8SetData vb, 0, size * count, 0, verts(0)
        
    End If
    

End Sub

Public Sub GetTransformedBox()
g_d3dx.ComputeBoundingBoxFromMesh TransformedMesh, transminExt, transmaxExt
End Sub
'-----------------------------------------------------------------------------
' Name: GetLocalBox
' Desc: Returns the extent of the mesh in the local coordinate system
'----------------------------------------------------------------------------
Public Sub GetLocalBox()
    Dim v As D3DVECTOR
    Dim heightV As D3DVECTOR
    Dim depthV As D3DVECTOR
    Dim width As Single
    Dim height As Single
    Dim depth As Single
    Dim adjBuf As D3DXBuffer
    
    g_d3dx.ComputeBoundingBoxFromMesh mesh, minExt, maxExt
        
    ReDim BoundBoxVertex(7)
    BoundBoxVertex(0) = vec3(maxExt.X, maxExt.Y, maxExt.z)
    BoundBoxVertex(1) = vec3(maxExt.X, maxExt.Y, minExt.z)
    BoundBoxVertex(2) = vec3(maxExt.X, minExt.Y, maxExt.z)
    BoundBoxVertex(3) = vec3(maxExt.X, minExt.Y, minExt.z)
    BoundBoxVertex(4) = vec3(minExt.X, maxExt.Y, maxExt.z)
    BoundBoxVertex(5) = vec3(minExt.X, maxExt.Y, minExt.z)
    BoundBoxVertex(6) = vec3(minExt.X, minExt.Y, maxExt.z)
    BoundBoxVertex(7) = vec3(minExt.X, minExt.Y, minExt.z)
    
'    D3DXVec3Subtract V, maxExt, minExt
'    Set BoundBox = g_d3dx.CreateMeshFVF(6, 8, D3DXMESH_POINTS, D3DFVF_XYZ, g_dev)
    
'    Set BoundBoxVB = BoundBox.GetVertexBuffer
'    D3DVertexBuffer8SetData BoundBoxVB, 0, Len(BoundBoxVertex(0)) * 8, 0, BoundBoxVertex(0)
   
End Sub

'-----------------------------------------------------------------------------
' Name: Destroy
' Desc: release any reference to frame and texture objects
'-----------------------------------------------------------------------------
Sub Destroy()
    
    'Releases all objects (does leave 1 element in the array)
    ReDim m_Textures(0)
    ReDim m_Materials(0)
    ReDim m_TextureNames(0)
    ReDim m_BoneFrames(0)
    ReDim m_BoneOffsetMat(0)
    ReDim m_BoneCombos(0)
    m_NumMaterials = 0
    bUseMaterials = True
    Set mesh = Nothing
    Set skinmesh = Nothing
End Sub


'-----------------------------------------------------------------------------
' Name: ComputeBoundingVolumes
' Desc: Makes BoundingSphere valid
'-----------------------------------------------------------------------------
Public Sub ComputeBoundingVolumes()
    g_d3dx.ComputeBoundingSphereFromMesh mesh, m_SphereCenter, SphereRadius
    bHasSphere = True
End Sub


'-----------------------------------------------------------------------------
' Name: RenderEx
' Desc: Render Mesh
' Params:
'   dev                         the device to draw to
'   bDrawOpaqueSubsets          draws all triangles that do not have alpha
'   bDrawOpaqueSubsets          draws all triangles that have alpha
'                               (note Blending renderstates are modified)
'
' Note: do not use for skinned meshes
'-----------------------------------------------------------------------------
'Sub RenderEx(dev As Direct3DDevice8, bDrawOpaqueSubsets As Boolean, bDrawAlphaSubsets As Boolean)
'    If mesh Is Nothing Then Exit Sub
'    Dim i As Long
    
'    'If bSkinned = True Then Exit Sub

'    ' Frist, draw the subsets without alpha
'    If (bDrawOpaqueSubsets) Then
'        For i = 0 To m_NumMaterials - 1
        
                
        
'            If (bUseMaterials) Then
'                If m_Materials(i).diffuse.a = 1# Then
'                    g_dev.SetMaterial m_Materials(i)
                    
'                    If g_bDontDrawTextures Then
'                        g_dev.SetTexture 0, Nothing
'                    Else
'                        g_dev.SetTexture 0, m_Textures(i)
'                    End If
                    
'                    mesh.DrawSubset i
'                End If
'            ElseIf (bUseMaterialOverride) Then
'                If m_MaterialOverride.diffuse.a = 1# Then
'                    If g_bDontDrawTextures Then
'                        g_dev.SetTexture 0, Nothing
'                    Else
'                        g_dev.SetTexture 0, m_Textures(i)
'                    End If
'                    g_dev.SetMaterial m_MaterialOverride
'                    mesh.DrawSubset i
'                End If
'            Else
'                mesh.DrawSubset i
'            End If
                    
'        Next
'    End If

    
    ' Then, draw the subsets with alpha
'    If (bDrawAlphaSubsets And (bUseMaterialOverride Or bUseMaterials)) Then
'        For i = 0 To m_NumMaterials - 1
        
'            If (bUseMaterials) Then
'                    If (m_Materials(i).diffuse.a < 1#) Then
'                          g_dev.SetMaterial m_Materials(i)
'                          g_dev.SetTexture 0, m_Textures(i)
'                          mesh.DrawSubset i
'                    End If
'            ElseIf (bUseMaterialOverride) Then
'                    If (m_MaterialOverride.diffuse.a < 1#) Then
'                          g_dev.SetMaterial m_MaterialOverride
'                          g_dev.SetTexture 0, m_Textures(i)
'                          mesh.DrawSubset i
'                    End If
'            End If
'        Next
'    End If
    
'End Sub

'-----------------------------------------------------------------------------
' Name: Render
' Desc: Render the mesh to the given device
'
' Note: Do not use for skinned meshes
'
'-----------------------------------------------------------------------------
Sub Render(dev As Direct3DDevice8)
    Dim i As Long
    Dim minExt As D3DVECTOR
    Dim maxExt As D3DVECTOR
    
    If mesh Is Nothing Then Exit Sub
    
    If bSkinned = True Then Exit Sub
    
    
    Transform
    If (bUseMaterials) Then
        For i = 0 To m_NumMaterials - 1
            g_dev.SetMaterial m_Materials(i)
            g_dev.SetTexture i, m_Textures(i)
            mesh.DrawSubset i
        Next
    Else
        For i = 0 To m_NumMaterials - 1
            mesh.DrawSubset i
        Next
    End If
    
    If curmesh = meshnum And curmesh <> 0 Then
'        TransformAbsolute
'        GetLocalBox
        DrawBoundBox
    End If
End Sub

Private Sub DrawBoundBox()
Dim devmat As D3DMATERIAL8
Dim a As Long

GetLocalBox

'a = g_dev.GetRenderState(D3DRS_AMBIENT)

devmat.Ambient.r = 1
devmat.Ambient.g = 1
'g_dev.SetRenderState D3DRS_AMBIENT, &HFFFFFFF
g_dev.SetMaterial devmat
g_dev.DrawIndexedPrimitiveUP D3DPT_LINELIST, 0, 8, 12, BoundBoxIndex(0), D3DFMT_INDEX16, BoundBoxVertex(0), Len(BoundBoxVertex(0))
'g_dev.SetRenderState D3DRS_AMBIENT, a
End Sub
'-----------------------------------------------------------------------------
' Name: RenderSkin
' Desc: Render the Mesh as skin
' Note: The mesh must have been loaded as a skin and bones must have been attached
'-----------------------------------------------------------------------------
'Sub RenderSkin()

'    If Not bSkinned Then Exit Sub
    
'    Dim ipAttr As Long  'bonecombo attribute
'    Dim matId As Long   'matrix id
    
'    Dim i As Long
'    Dim mat2 As D3DMATRIX
'    Dim mat1 As D3DMATRIX
'    Dim mat0 As D3DMATRIX

    
'    g_dev.SetRenderState D3DRS_VERTEXBLEND, m_maxFaceInfl - 1
        
'    For ipAttr = 0 To m_NumBoneCombos - 1
'        For i = 0 To m_maxFaceInfl - 1
'                matId = m_BoneComboIds(ipAttr).List(i)
                
                'If we get a MatId of -1 then all the vertex weights are 0
                'and we dont need to set the transform for this bone
'                If matId <> -1 Then
'                    mat0 = m_BoneFrames(matId).GetUpdatedMatrix()
'                    mat1 = m_BoneOffsetMat(matId)
'                    D3DXMatrixMultiply mat2, mat1, mat0
'                    g_dev.SetTransform m_worldMatHandles(i), mat2
'                End If
'        Next
'        g_dev.SetTexture 0, m_Textures(m_BoneCombos(ipAttr).AttribId)
'        g_dev.SetMaterial m_Materials(m_BoneCombos(ipAttr).AttribId)
        
'        mesh.DrawSubset ipAttr
        
'    Next
    
'    g_dev.SetRenderState D3DRS_VERTEXBLEND, 0
    
    
'End Sub

'-----------------------------------------------------------------------------
' Name: GetMaterialCount
'
'---------------------------------------------------------------------------
Public Function GetMaterialCount() As Long
    GetMaterialCount = m_NumMaterials
End Function


'-----------------------------------------------------------------------------
' Name: SetMaterialOverride
' Desc: Sets the materail to be used in place of the ones loaded from file
' Note: to disable set bUseMaterialOverride to false
'-----------------------------------------------------------------------------
Public Sub SetMaterialOverride(m As D3DMATERIAL8)
    m_MaterialOverride = m
    bUseMaterialOverride = True
End Sub

'-----------------------------------------------------------------------------
' Name: GetMaterialOverride
' Desc:
'-----------------------------------------------------------------------------
Public Sub GetMaterialOverride(m As D3DMATERIAL8)
      m = m_MaterialOverride
End Sub


'-----------------------------------------------------------------------------
' Name: ClassName
' Desc:
'-----------------------------------------------------------------------------
Public Function ClassName() As String
    ClassName = "CD3DMesh"
End Function


'-----------------------------------------------------------------------------
' Name: InvalidateDeviceObjects
' Desc: Release reference to device dependent objects
'-----------------------------------------------------------------------------
Public Sub InvalidateDeviceObjects()
   'all framework objects are managed so nothing to do here
End Sub



'-----------------------------------------------------------------------------
' Name: RestoreDeviceObjects
' Desc: If we had any video memory objects they would need
'       to be reloaded here
'-----------------------------------------------------------------------------
Public Sub RestoreDeviceObjects(dev As Direct3DDevice8)
    
End Sub


'-----------------------------------------------------------------------------
' Name: InitFromD3DXMesh
' Desc: Allow mesh objects to be created from external D3DXMesh objects
'
'-----------------------------------------------------------------------------
Sub InitFromD3DXMesh(d3dxmeshIn As D3DXMesh)
'    Destroy
'    TranslationVector = vec3(0, 0, 0)
'    RotationVector = vec3(0, 0, 0)
    ScalingVector = vec3(1, 1, 1)
    bUseMaterials = False
    ReDim m_Materials(1)
    ReDim m_Textures(1)
    m_NumMaterials = 1
    Set mesh = d3dxmeshIn
End Sub


'-----------------------------------------------------------------------------
' Name: SetMaterialCount
' Desc: If a mesh was initialized with InitFromD3DXMesh
'       This function can allocate space for Materials and Textures
'-----------------------------------------------------------------------------
Sub SetMaterialCount(n As Long)
    m_NumMaterials = n
    ReDim Preserve m_Materials(n)
    ReDim Preserve m_Textures(n)
End Sub

'-----------------------------------------------------------------------------
' Name: SetMaterialTexture
' Desc: Sets the texture for a given material subset
' Note: use nothing to remove a texture
'-----------------------------------------------------------------------------
Sub SetMaterialTexture(n As Long, tex As Direct3DTexture8)
    Set m_Textures(n) = tex
End Sub

'-----------------------------------------------------------------------------
' Name: GetMaterialTexture
' Desc: returns a given texture for a material subset
'-----------------------------------------------------------------------------
Function GetMaterialTexture(n As Long) As Direct3DTexture8
    Set GetMaterialTexture = m_Textures(n)
End Function


'-----------------------------------------------------------------------------
' Name: SetMaterial
' Desc: Sets the material properties for a given material subset
'-----------------------------------------------------------------------------
Sub SetMaterial(n As Long, material As D3DMATERIAL8)
    m_Materials(n) = material
End Sub


'-----------------------------------------------------------------------------
' Name: GetMaterial
' Desc: returns material properties for a material subset
'-----------------------------------------------------------------------------
Function GetMaterial(n As Long) As D3DMATERIAL8
    GetMaterial = m_Materials(n)
End Function

Function Transform()
Dim matScale As D3DMATRIX
Dim matRotate As D3DMATRIX
Dim matTranslate As D3DMATRIX
Dim temp As D3DMATRIX
Dim matWorld As D3DMATRIX

'D3DXMatrixAffineTransformation matWorld, ScalingVector.X, vRotationCenter, RotationVector, TranslationVector
'D3DXMatrixTransformation matWorld, vRotationCenter, RotationVector, ScalingVector, vRotationCenter, RotationVector, TranslationVector
D3DXMatrixScaling matScale, ScalingVector.X, ScalingVector.Y, ScalingVector.z
D3DXMatrixRotationYawPitchRoll matRotate, RotationVector.Y, RotationVector.X, RotationVector.z
D3DXMatrixTranslation matTranslate, TranslationVector.X, TranslationVector.Y, TranslationVector.z
D3DXMatrixMultiply temp, matScale, matRotate
D3DXMatrixMultiply matWorld, temp, matTranslate
g_dev.SetTransform D3DTS_WORLD, matWorld
End Function

Private Sub ScaleIt(Nodes() As CUSTOMVERTEX)
Dim i As Integer

For i = 0 To UBound(Nodes)
'    Nodes(i).x = PrimalNodes(i).x + PrimalNodes(i).x * Scaling.x
'    Nodes(i).y = PrimalNodes(i).y + PrimalNodes(i).y * Scaling.y
'    Nodes(i).Z = PrimalNodes(i).Z + PrimalNodes(i).Z * Scaling.Z
    Nodes(i).X = Nodes(i).X * ScalingVector.X
    Nodes(i).Y = Nodes(i).Y * ScalingVector.Y
    Nodes(i).z = Nodes(i).z * ScalingVector.z
Next
End Sub
Private Sub RotateIt(Nodes() As CUSTOMVERTEX)
Dim i As Integer
Dim tempX As Double
Dim tempy As Double
Dim tempz As Double
Dim rotangle As Double

rotangle = RotationVector.X

For i = 0 To UBound(Nodes)
    tempX = Nodes(i).X
    tempy = Nodes(i).Y
    tempz = Nodes(i).z
    Nodes(i).Y = tempy * Cos(rotangle) - tempz * Sin(rotangle)
    Nodes(i).z = tempy * Sin(rotangle) + tempz * Cos(rotangle)
Next

rotangle = RotationVector.Y

For i = 0 To UBound(Nodes)
    tempX = Nodes(i).X
    tempy = Nodes(i).Y
    tempz = Nodes(i).z
    Nodes(i).X = tempX * Cos(rotangle) + tempz * Sin(rotangle)
    Nodes(i).z = -tempX * Sin(rotangle) + tempz * Cos(rotangle)
Next
rotangle = RotationVector.z

For i = 0 To UBound(Nodes)
    tempX = Nodes(i).X
    tempy = Nodes(i).Y
    tempz = Nodes(i).z
    Nodes(i).X = tempX * Cos(rotangle) - tempy * Sin(rotangle)
    Nodes(i).Y = tempX * Sin(rotangle) + tempy * Cos(rotangle)
Next
End Sub
Private Sub TranslateIt(Nodes() As CUSTOMVERTEX)
Dim i As Integer

For i = 0 To UBound(Nodes)
    Nodes(i).X = Nodes(i).X + TranslationVector.X
    Nodes(i).Y = Nodes(i).Y + TranslationVector.Y
    Nodes(i).z = Nodes(i).z + TranslationVector.z
Next
'MoveRefPoint Translation
End Sub

Public Function TransformAbsolute()
Dim g_VB As Direct3DVertexBuffer8
'Dim Vertices() As D3DVECTOR
Dim desc As D3DVERTEXBUFFER_DESC

Set TransformedMesh = mesh.CloneMeshFVF(D3DXMESH_VB_DYNAMIC, D3DFVF_CUSTOMVERTEX, g_dev)
'mesh.GetDeclaration decl
'Set TransformedMesh = mesh.CloneMesh(D3DXMESH_VB_DYNAMIC, decl, g_dev)

ReDim Vertex(0 To TransformedMesh.GetNumVertices - 1)
Set g_VB = TransformedMesh.GetVertexBuffer
'g_VB.GetDesc desc
'D3DVertexBuffer8GetData g_VB, 0, Len(Vertices(0)) * TransformedMesh.GetNumVertices, 0, Vertices(0)
'D3DVertexBuffer8GetData g_VB, 0, desc.size, 0, Vertices(0)
D3DVertexBuffer8GetData g_VB, 0, Len(Vertex(0)) * TransformedMesh.GetNumVertices, 0, Vertex(0)

ScaleIt Vertex()
RotateIt Vertex()
TranslateIt Vertex()

'D3DVertexBuffer8SetData g_VB, 0, Len(Vertices(0)) * TransformedMesh.GetNumVertices, 0, Vertices(0)
'D3DVertexBuffer8SetData g_VB, 0, desc.size, 0, Vertices(0)
D3DVertexBuffer8SetData g_VB, 0, Len(Vertex(0)) * TransformedMesh.GetNumVertices, 0, Vertex(0)

'g_d3dx.ComputeNormals TransformedMesh
'GetLocalBox
End Function

Public Sub TranslateX(Value As Single)
TranslationVector.X = TranslationVector.X + Value
End Sub
Public Sub TranslateY(Value As Single)
TranslationVector.Y = TranslationVector.Y + Value
End Sub
Public Sub TranslateZ(Value As Single)
TranslationVector.z = TranslationVector.z + Value
End Sub

Public Sub RotateX(Value As Single)
RotationVector.X = RotationVector.X + Value
End Sub
Public Sub RotateY(Value As Single)
RotationVector.Y = RotationVector.Y + Value
End Sub
Public Sub RotateZ(Value As Single)
RotationVector.z = RotationVector.z + Value
End Sub

Public Sub ScaleUniform(Value As Single)
ScalingVector.X = ScalingVector.X + Value
ScalingVector.Y = ScalingVector.Y + Value
ScalingVector.z = ScalingVector.z + Value
End Sub

Public Function IntersectRayBound(RayStart As D3DVECTOR, RayDir As D3DVECTOR) As Boolean
GetTransformedBox
IntersectRayBound = g_d3dx.BoxBoundProbe(transminExt, transmaxExt, RayStart, RayDir)
End Function

Public Sub GetMeshVertex()
Dim g_VB As Direct3DVertexBuffer8

'Set TransformedMesh = mesh.CloneMeshFVF(D3DXMESH_VB_DYNAMIC, D3DFVF_XYZ, g_dev)

ReDim Vertex(0 To TransformedMesh.GetNumVertices - 1)
Set g_VB = TransformedMesh.GetVertexBuffer
D3DVertexBuffer8GetData g_VB, 0, Len(Vertex(0)) * TransformedMesh.GetNumVertices, 0, Vertex(0)
End Sub

Public Sub GetMeshIndex()
Dim g_IB As Direct3DIndexBuffer8

'Set TransformedMesh = mesh.CloneMeshFVF(D3DXMESH_VB_DYNAMIC, D3DFVF_XYZ, g_dev)

ReDim Index(0 To TransformedMesh.GetNumFaces - 1)
Set g_IB = TransformedMesh.GetIndexBuffer
D3DIndexBuffer8GetData g_IB, 0, Len(Index(0)) * TransformedMesh.GetNumFaces, 0, Index(0)
End Sub

Public Function GetVertexNormal(VertexNum As Long) As D3DVECTOR
GetVertexNormal = Vertex(VertexNum).normal
End Function

Public Function GetIndex(IndexNum As Long, VertexNum As Integer) As Long
If VertexNum = 1 Then
    GetIndex = Index(IndexNum).v1
ElseIf VertexNum = 2 Then
    GetIndex = Index(IndexNum).v2
ElseIf VertexNum = 3 Then
    GetIndex = Index(IndexNum).v3
End If
End Function

