VERSION 5.00
Begin VB.Form CameraForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transform Camera"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Transformation Axis"
      Height          =   975
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   2415
      Begin VB.OptionButton Option5 
         Caption         =   "Z"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Y"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "X"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Camera Component"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "Focus Point"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Camera Point"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "CameraForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
TransformCamera = True
TransformMesh = False
TransformLight = False
End Sub

