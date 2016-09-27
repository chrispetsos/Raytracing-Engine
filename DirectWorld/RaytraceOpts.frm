VERSION 5.00
Begin VB.Form RaytraceOpts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raytrace Options"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   2235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "Raytrace Depth"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "Weight Treshold"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "RaytraceOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
RayTraceDepth = Val(Text1.Text)
End Sub

Private Sub Text2_Change()
WeightTreshold = Val(Text2.Text)
End Sub
