VERSION 5.00
Begin VB.Form RenderOpts 
   Caption         =   "AVIFile Tutorial BMP to AVI"
   ClientHeight    =   3585
   ClientLeft      =   1125
   ClientTop       =   2415
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.TextBox txtFPS 
      Height          =   285
      Left            =   4035
      TabIndex        =   6
      Text            =   "30"
      Top             =   2520
      Width           =   465
   End
   Begin VB.CommandButton cmdWriteAVI 
      Caption         =   "Write AVI..."
      Height          =   480
      Left            =   2475
      TabIndex        =   4
      Top             =   2925
      Width           =   2025
   End
   Begin VB.ListBox lstDIBList 
      Height          =   1230
      Left            =   180
      TabIndex        =   3
      Top             =   135
      Width           =   4290
   End
   Begin VB.CommandButton cmdFileOpen 
      Caption         =   "Add BMP file to list..."
      Height          =   480
      Left            =   180
      TabIndex        =   1
      Top             =   1530
      Width           =   2025
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear file list"
      Enabled         =   0   'False
      Height          =   480
      Left            =   2475
      TabIndex        =   0
      Top             =   1530
      Width           =   2025
   End
   Begin VB.Label lblfps 
      Caption         =   "Frames per second (1 - 30):"
      Height          =   195
      Left            =   1755
      TabIndex        =   5
      Top             =   2565
      Width           =   2040
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   180
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   1230
   End
   Begin VB.Label lblStatus 
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   2160
      Width           =   4290
   End
End
Attribute VB_Name = "RenderOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'simple UDT containing parameters of first BMP file user chooses
'all the following BMPs should be the same format so there will be no problems in writing the vidstream
Private Type PARAMS
    Init As Boolean
    Width As Long
    Height As Long
    bpp As Long
End Type

Private Declare Function SetRect Lib "user32.dll" _
    (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long 'BOOL

Private m_params As PARAMS





Private Sub Form_Unload(Cancel As Integer)
    Call AVIFileExit   '// releases AVIFile library
End Sub

Private Sub Form_Load()
'adds a bmp to list of files to create video stream from
Dim szFileName As String
Dim file As cFileDlg
Dim bmp As cDIB
Static InitDir As String

Call AVIFileInit   '// opens AVIFile library
'Set file dialog parameters
Set file = New cFileDlg
With file
    .DlgTitle = "Choose BMP file to add to video stream"
    .Filter = "BMP Files|*.bmp:DIB Files|*.dib"
    If InitDir <> "" Then
        file.InitDirectory = InitDir
    End If
End With
For i = 1 To 150
    szFileName = App.Path & "\test" & FixFiveDigit(CStr(i)) & ".bmp"
    'get filename from user
'    If file.VBGetOpenFileName(szFileName) = True Then
        Set bmp = New cDIB
        If bmp.CreateFromFile(szFileName) Then 'file is a valid BMP
            If m_params.Init Then 'this is not the first file - it must be the same format
                If (bmp.Width <> m_params.Width) _
                    Or (bmp.Height <> m_params.Height) _
                    Or (bmp.BitCount <> m_params.bpp) Then
                    MsgBox "Chosen bitmap file is a different format!", vbInformation, App.title 'format is wrong
                Else
                    imgPreview.Picture = LoadPicture(szFileName) 'format is OK -add file to list
                    lstDIBList.AddItem szFileName
                End If
            Else 'this is the first file in the list so save format info too
                With m_params
                    .Init = True
                    .Width = bmp.Width
                    .Height = bmp.Height
                    .bpp = bmp.BitCount
                End With
                imgPreview.Picture = LoadPicture(szFileName)
                lstDIBList.AddItem szFileName
            End If
            cmdClearList.Enabled = True 'make sure clear button is enabled
            cmdWriteAVI.Enabled = True 'allow user to call AVI write functions
        End If
        Set bmp = Nothing
    'End If
    'save last directory for user
'    InitDir = file.InitDirectory
    Set file = Nothing
Next
End Sub

Private Sub cmdClearList_Click()
'reset file list - unload picture - reset format params
lstDIBList.Clear
imgPreview.Picture = LoadPicture()
With m_params
    .bpp = 0
    .Height = 0
    .Width = 0
    .Init = False
End With
cmdClearList.Enabled = False
cmdWriteAVI.Enabled = False
End Sub

Public Sub cmdWriteAVI_Click()
 
    Dim file As cFileDlg
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim bmp As cDIB
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim BI As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long
    
   Debug.Print
    Set file = New cFileDlg
    'get an avi filename from user
    With file
        .DefaultExt = "avi"
        .DlgTitle = "Choose a filename to save AVI to..."
        .Filter = "AVI Files|*.avi"
        .OwnerHwnd = Me.hWnd
    End With
    szOutputAVIFile = "MyAVI.avi"
    If file.VBGetSaveFileName(szOutputAVIFile) <> True Then Exit Sub
        
'    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB
    lstDIBList.ListIndex = 0
    If bmp.CreateFromFile(lstDIBList.Text) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.title
        GoTo error
    End If

'   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = val(txtFPS)                        '// fps
        .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)       '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(Me.hWnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                        1, _
                        ps, _
                        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If
    
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

    RenderFrm.Picture1.Print "Creating AVI..."
    RenderFrm.Picture1.Refresh

'   Now write out each video frame
    For i = 0 To lstDIBList.ListCount - 1
        lstDIBList.ListIndex = i
        bmp.CreateFromFile (lstDIBList.Text) 'load the bitmap (ignore errors)
        res = AVIStreamWrite(psCompressed, _
                            i, _
                            1, _
                            bmp.PointerToBits, _
                            bmp.SizeImage, _
                            AVIIF_KEYFRAME, _
                            ByVal 0&, _
                            ByVal 0&)
        If res <> AVIERR_OK Then GoTo error
        'Show user feedback
        imgPreview.Picture = LoadPicture(lstDIBList.Text)
        imgPreview.Refresh
        lblStatus = "Frame number " & i & " saved"
        lblStatus.Refresh
    Next
    lblStatus = "Finished!"

error:
'   Now close the file
    Set file = Nothing
    Set bmp = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit

    If (res <> AVIERR_OK) Then
        MsgBox "There was an error writing the file.", vbInformation, App.title
    End If
End Sub

