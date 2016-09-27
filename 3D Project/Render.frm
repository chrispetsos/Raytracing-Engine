VERSION 5.00
Begin VB.Form RenderFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Render"
   ClientHeight    =   10020
   ClientLeft      =   630
   ClientTop       =   630
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   12300
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H8000000E&
      Height          =   9975
      Left            =   0
      ScaleHeight     =   662.341
      ScaleMode       =   0  'User
      ScaleWidth      =   814.649
      TabIndex        =   0
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "RenderFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
With MainForm
    .theWorld.VertexNormals = .Check1.Value
    .theWorld.FaceNormals = .Check2.Value
    .theWorld.Mesh = .Check3.Value
    .theWorld.Flat = .Option7.Value
    .theWorld.Gouraud = .Option8.Value
    .theWorld.Phong = False
End With

End Sub

Private Sub Form_Terminate()
With MainForm
    .theWorld.VertexNormals = .Check1.Value
    .theWorld.FaceNormals = .Check2.Value
    .theWorld.Mesh = .Check3.Value
    .theWorld.Flat = .Option7.Value
    .theWorld.Gouraud = .Option8.Value
    .theWorld.Phong = False
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
With MainForm
    .theWorld.VertexNormals = .Check1.Value
    .theWorld.FaceNormals = .Check2.Value
    .theWorld.Mesh = .Check3.Value
    .theWorld.Flat = .Option7.Value
    .theWorld.Gouraud = .Option8.Value
    .theWorld.Phong = False
End With

End Sub
