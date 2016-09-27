VERSION 5.00
Begin VB.Form AddNewObject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Object"
   ClientHeight    =   1770
   ClientLeft      =   5880
   ClientTop       =   5040
   ClientWidth     =   1755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   1755
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE OBJECT"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "255"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Color B"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Color G"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Color R"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "AddNewObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub
