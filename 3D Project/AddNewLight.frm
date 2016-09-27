VERSION 5.00
Begin VB.Form AddNewLightFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Light"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "0.5"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "0.1"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "500"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "30"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "60"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE LIGHT"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Ambient Constant (Ia*Ka)"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Light Source Intensity (Ii)"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Light Point Rho"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Light Point Theta"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Light Point Phi"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "AddNewLightFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim tempLight As New Light

tempLight.Create val(Text1.Text), val(Text2.Text), val(Text3.Text), val(Text4.Text), val(Text5.Text)
MainForm.theWorld.AddLight tempLight
Me.Hide
With MainForm
    .Label10.Caption = .theWorld.ActiveLight.Rho
    .Label11.Caption = .theWorld.ActiveLight.Theta
    .Label12.Caption = .theWorld.ActiveLight.Phi
    .Label16(0).Caption = .theWorld.ActiveLight.Ambient
    .Label16(1).Caption = .theWorld.ActiveLight.Intensity
End With

End Sub
