VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LightForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Light Form"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Cast Shadows"
      Height          =   255
      Left            =   2520
      TabIndex        =   33
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply Material"
      Height          =   375
      Left            =   3600
      TabIndex        =   32
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   0
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   1
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Index           =   2
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   7
      Text            =   "1"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   6
      Text            =   "1"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   5
      Text            =   "1"
      Top             =   3360
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Z"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Y"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Current Light"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Ambient"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   30
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   29
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   28
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   27
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   26
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   25
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Diffuse"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   23
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   22
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   21
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Specular"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   16
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   15
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "A"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   14
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Translation"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "LightForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurMat As D3DMATERIAL8

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lights(curLight).SetState Check1.Value
End Sub

Private Sub Check2_Click()
Lights(curLight).CastShadows = Check2.Value
End Sub

Private Sub Command1_Click()
'k = g_dev.GetRenderState(D3DRS_AMBIENT)

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
Lights(curLight).SetMaterial CurMat
g_dev.SetLight curLight - 1, Lights(curLight).GetD3DLight 'let d3d know about the light
End Sub

Private Sub Form_Activate()
TransformLight = True
TransformCamera = False
TransformMesh = False

Check1.Value = Lights(curLight).GetState
Check2.Value = Lights(curLight).CastShadows
CurMat = Lights(curLight).GetMaterial
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
