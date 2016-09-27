VERSION 5.00
Begin VB.Form TransformForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transformations"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Y"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Z"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option4 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Y"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Z"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Scale Uniform"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Translation"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Rotation"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "TransformForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
TransformMesh = True
TransformCamera = False
TransformLight = False
End Sub

