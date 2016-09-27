VERSION 5.00
Begin VB.Form RaytraceForm 
   Caption         =   "RayTrace"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "RaytraceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.width = MainForm.width
Me.height = MainForm.height
Picture1.width = MainForm.width
Picture1.height = MainForm.height

End Sub
