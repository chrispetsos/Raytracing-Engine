VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MaterialForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Editor"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   52
      Top             =   5280
      Width           =   2535
   End
   Begin VB.HScrollBar HScroll2 
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      LargeChange     =   5
      Left            =   2280
      Max             =   1000
      TabIndex        =   45
      Top             =   4800
      Value           =   500
      Width           =   3735
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Index           =   10
      LargeChange     =   5
      Left            =   2280
      Max             =   1000
      TabIndex        =   44
      Top             =   4080
      Value           =   500
      Width           =   3735
   End
   Begin VB.HScrollBar HScroll2 
      Enabled         =   0   'False
      Height          =   255
      Index           =   11
      LargeChange     =   5
      Left            =   2280
      Max             =   2000
      Min             =   1000
      TabIndex        =   43
      Top             =   4440
      Value           =   1000
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply Material"
      Height          =   375
      Left            =   5280
      TabIndex        =   42
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   41
      Text            =   "1"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   40
      Text            =   "1"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   39
      Text            =   "1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   5400
      TabIndex        =   38
      Text            =   "1"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   37
      Text            =   "0"
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   3
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   32
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   2
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   1
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   0
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Ktg(Global Transmission)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   51
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "0.5"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   6120
      TabIndex        =   50
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Krg (Global Reflection)"
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   49
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "0.5"
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   48
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "h (relative index of refraction)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   11
      Left            =   -120
      TabIndex        =   47
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label16 
      Caption         =   "10"
      Enabled         =   0   'False
      Height          =   255
      Index           =   11
      Left            =   6120
      TabIndex        =   46
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Power"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   35
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   34
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   33
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   31
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   30
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   29
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "R"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Emmisive"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   26
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   25
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   24
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   22
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   21
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   20
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "R"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Specular"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   12
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "R"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Diffuse"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   4
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "R"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Ambient"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "MaterialForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurMat As D3DMATERIAL8

Private Sub Command1_Click()
If curmesh <> 0 Then
    CurMat.Ambient.r = Val(Label6(0).Caption) / 255
    CurMat.Ambient.g = Val(Label7(0).Caption) / 255
    CurMat.Ambient.b = Val(Label8(0).Caption) / 255
    CurMat.Ambient.a = Val(Text2(0).Text)
    CurMat.diffuse.r = Val(Label6(1).Caption) / 255
    CurMat.diffuse.g = Val(Label7(1).Caption) / 255
    CurMat.diffuse.b = Val(Label8(1).Caption) / 255
    CurMat.diffuse.a = Val(Text2(1).Text)
    CurMat.specular.r = Val(Label6(2).Caption) / 255
    CurMat.specular.g = Val(Label7(2).Caption) / 255
    CurMat.specular.b = Val(Label8(2).Caption) / 255
    CurMat.specular.a = Val(Text2(2).Text)
    CurMat.emissive.r = Val(Label6(3).Caption) / 255
    CurMat.emissive.g = Val(Label7(3).Caption) / 255
    CurMat.emissive.b = Val(Label8(3).Caption) / 255
    CurMat.emissive.a = Val(Text2(3).Text)
    CurMat.power = Val(Text1.Text)
    g_Mesh(curmesh).SetMaterial 0, CurMat
'    g_Mesh(curmesh).Kd = Val(HScroll2(2).Value) / 1000
'    g_Mesh(curmesh).Ks = Val(HScroll2(3).Value) / 1000
'    g_Mesh(curmesh).n = Val(HScroll2(4).Value)
'    g_Mesh(curmesh).Kl = Val(HScroll2(8).Value) / 1000
    g_Mesh(curmesh).Ktg = Val(HScroll2(9).Value) / 1000
    g_Mesh(curmesh).Krg = Val(HScroll2(10).Value) / 1000
    g_Mesh(curmesh).h = Val(HScroll2(11).Value) / 1000
End If
End Sub


Private Sub Command2_Click()
Dim tex As Direct3DTexture8

CommonDialog1.ShowOpen
'Set tex = g_d3dx.CreateTextureFromFile(g_dev, CommonDialog1.FileName)
Set tex = g_d3dx.CreateTextureFromFileEx(g_dev, CommonDialog1.FileName, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, ByVal 0, ByVal 0)
g_dev.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
g_dev.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
g_dev.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
g_dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
g_Mesh(curmesh).SetMaterialTexture 0, tex
End Sub

Private Sub Form_Activate()
If curmesh <> 0 Then
    CurMat = g_Mesh(curmesh).GetMaterial(0)
    Label6(0).Caption = CStr(CurMat.Ambient.r * 255)
    Label7(0).Caption = CStr(CurMat.Ambient.g * 255)
    Label8(0).Caption = CStr(CurMat.Ambient.b * 255)
    Text2(0).Text = CStr(CurMat.Ambient.a)
    Picture1(0).BackColor = RGB(Val(Label6(0).Caption), Val(Label7(0).Caption), Val(Label8(0).Caption))
    Label6(1).Caption = CStr(CurMat.diffuse.r * 255)
    Label7(1).Caption = CStr(CurMat.diffuse.g * 255)
    Label8(1).Caption = CStr(CurMat.diffuse.b * 255)
    Text2(1).Text = CStr(CurMat.diffuse.a)
    Picture1(1).BackColor = RGB(Val(Label6(1).Caption), Val(Label7(1).Caption), Val(Label8(1).Caption))
    Label6(2).Caption = CStr(CurMat.specular.r * 255)
    Label7(2).Caption = CStr(CurMat.specular.g * 255)
    Label8(2).Caption = CStr(CurMat.specular.b * 255)
    Text2(2).Text = CStr(CurMat.specular.a)
    Picture1(2).BackColor = RGB(Val(Label6(2).Caption), Val(Label7(2).Caption), Val(Label8(2).Caption))
    Label6(3).Caption = CStr(CurMat.emissive.r * 255)
    Label7(3).Caption = CStr(CurMat.emissive.g * 255)
    Label8(3).Caption = CStr(CurMat.emissive.b * 255)
    Text2(3).Text = CStr(CurMat.emissive.a)
    Picture1(3).BackColor = RGB(Val(Label6(3).Caption), Val(Label7(3).Caption), Val(Label8(3).Caption))
    Text1.Text = CStr(CurMat.power)
    
'    HScroll2(2).Value = g_Mesh(curmesh).Kd * 1000
'    HScroll2(3).Value = g_Mesh(curmesh).Ks * 1000
'    HScroll2(4).Value = g_Mesh(curmesh).n
'    HScroll2(8).Value = g_Mesh(curmesh).Kl * 1000
    HScroll2(9).Value = g_Mesh(curmesh).Ktg * 1000
    HScroll2(10).Value = g_Mesh(curmesh).Krg * 1000
    HScroll2(11).Value = g_Mesh(curmesh).h * 1000
End If
End Sub

Private Sub HScroll2_Change(Index As Integer)
If Index = 0 Then
    Label16(Index) = HScroll2(Index).Value
'    theWorld.ActiveLight.Ambient = HScroll2(Index).Value / 1000
ElseIf Index = 1 Then
    Label16(Index) = HScroll2(Index).Value
'    theWorld.ActiveLight.Intensity = HScroll2(Index).Value / 1000
ElseIf Index = 2 Then
    Label16(Index) = HScroll2(Index).Value / 1000
'    theWorld.ActiveObject.Kd = HScroll2(Index).Value / 1000
ElseIf Index = 3 Then
    Label16(Index) = HScroll2(Index).Value / 1000
'    theWorld.ActiveObject.Ks = HScroll2(Index).Value / 1000
ElseIf Index = 4 Then
    Label16(Index) = HScroll2(Index).Value
'    theWorld.ActiveObject.n = HScroll2(Index).Value
ElseIf Index = 5 Then
    Label16(Index) = HScroll2(Index).Value
'    theWorld.ActiveObject.hue = HScroll2(Index).Value
ElseIf Index = 6 Then
    Label16(Index) = HScroll2(Index).Value
'    theWorld.ActiveObject.Saturation = HScroll2(Index).Value / 1000
ElseIf Index = 7 Then
    Label16(Index) = HScroll2(Index).Value
'    theWorld.ActiveObject.Lightness = HScroll2(Index).Value / 1000
ElseIf Index = 8 Then
    Label16(Index) = HScroll2(Index).Value / 1000
'    theWorld.ActiveObject.Kl = HScroll2(Index).Value / 1000
ElseIf Index = 9 Then
    Label16(Index) = HScroll2(Index).Value / 1000
'    theWorld.ActiveObject.Ktg = HScroll2(Index).Value / 1000
ElseIf Index = 10 Then
    Label16(Index) = HScroll2(Index).Value / 1000
'    theWorld.ActiveObject.Krg = HScroll2(Index).Value / 1000
ElseIf Index = 11 Then
    Label16(Index) = HScroll2(Index).Value / 1000
'    theWorld.ActiveObject.hta = HScroll2(Index).Value / 1000
End If
End Sub

Private Sub Picture1_Click(Index As Integer)
Dim b As D3DCOLORVALUE

CommonDialog1.ShowColor
b = LONGtoD3DCOLORVALUE(CommonDialog1.color)
Label6(Index).Caption = CStr(b.r * 255)
Label7(Index).Caption = CStr(b.g * 255)
Label8(Index).Caption = CStr(b.b * 255)
'Label9(Index).Caption = CStr(b.a * 255)
Picture1(Index).BackColor = CommonDialog1.color
End Sub
