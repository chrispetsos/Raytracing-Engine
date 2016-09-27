VERSION 5.00
Begin VB.Form AddNewCamera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Camera"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE CAMERA"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "1000"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "1000"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "1000"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Focus Point Z"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Focus Point Y"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Focus Point X"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Camera Point Z"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Camera Point Y"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Camera Point X"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "AddNewCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim CPoint As New Vector3D
Dim FPoint As New Vector3D
Dim tempCam As New Camera

CPoint.x = val(Text1.Text)
CPoint.y = val(Text2.Text)
CPoint.Z = val(Text3.Text)
FPoint.x = val(Text4.Text)
FPoint.y = val(Text5.Text)
FPoint.Z = val(Text6.Text)
tempCam.Create CPoint, FPoint, 0
MainForm.theWorld.AddCamera tempCam
Me.Hide
With MainForm
    .theWorld.Raster .Check1.Value, .Check2.Value, .Check3.Value, .Option7.Value, .Option8.Value, .Option10.Value, True, True
    .Label4.Caption = .theWorld.ActiveCamera.CameraPoint.x
    .Label5.Caption = .theWorld.ActiveCamera.CameraPoint.y
    .Label6.Caption = .theWorld.ActiveCamera.CameraPoint.Z
    .Label17.Caption = .theWorld.ActiveCamera.FocusPoint.x
    .Label14.Caption = .theWorld.ActiveCamera.FocusPoint.y
    .Label13.Caption = .theWorld.ActiveCamera.FocusPoint.Z
    .Label21.Caption = .theWorld.ActiveCamera.TwistAngle
End With
End Sub

