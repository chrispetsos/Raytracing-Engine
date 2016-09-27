VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame13 
      Caption         =   "Camera Animation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   104
      Top             =   9840
      Width           =   7575
      Begin VB.CommandButton Command36 
         Caption         =   "RENDER"
         Height          =   435
         Left            =   6600
         TabIndex        =   113
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command35 
         Caption         =   "End"
         Height          =   255
         Left            =   5520
         TabIndex        =   112
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Start"
         Height          =   255
         Left            =   2280
         TabIndex        =   111
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4800
         TabIndex        =   110
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   109
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   108
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   107
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   106
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "OBJECT PROPERTIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   840
      TabIndex        =   85
      Top             =   7680
      Width           =   6615
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   7
         LargeChange     =   5
         Left            =   1920
         Max             =   1000
         TabIndex        =   101
         Top             =   720
         Value           =   500
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   6
         LargeChange     =   5
         Left            =   1920
         Max             =   1000
         TabIndex        =   98
         Top             =   1080
         Value           =   500
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   5
         LargeChange     =   5
         Left            =   1920
         Max             =   360
         TabIndex        =   95
         Top             =   360
         Value           =   360
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   2
         LargeChange     =   5
         Left            =   1920
         Max             =   1000
         TabIndex        =   88
         Top             =   1440
         Value           =   500
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   3
         LargeChange     =   5
         Left            =   1920
         Max             =   1000
         TabIndex        =   87
         Top             =   1800
         Value           =   500
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   4
         LargeChange     =   50
         Left            =   1920
         Max             =   500
         SmallChange     =   5
         TabIndex        =   86
         Top             =   2160
         Value           =   10
         Width           =   3735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Lightness"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   103
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "0.5"
         Height          =   255
         Index           =   7
         Left            =   5760
         TabIndex        =   102
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "0.5"
         Height          =   255
         Index           =   6
         Left            =   5760
         TabIndex        =   100
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Saturation"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   99
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "360"
         Height          =   255
         Index           =   5
         Left            =   5760
         TabIndex        =   97
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Hue"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   96
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Kd (Diffuse Light)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   94
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "0.5"
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   93
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Ks (Specular Light)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   92
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "0.5"
         Height          =   255
         Index           =   3
         Left            =   5760
         TabIndex        =   91
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "n (Spread of Specular)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   90
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "10"
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   89
         Top             =   2160
         Width           =   615
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "LIGHT SOURCE PROPERTIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7560
      TabIndex        =   78
      Top             =   8400
      Width           =   6615
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   1
         LargeChange     =   5
         Left            =   2040
         Max             =   1000
         TabIndex        =   82
         Top             =   840
         Value           =   100
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Index           =   0
         LargeChange     =   5
         Left            =   2040
         Max             =   1000
         TabIndex        =   79
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label16 
         Caption         =   "0.01"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   84
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Light Source Intensity (Ii)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   83
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   81
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Ambient Constant (Ia*Ka)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Object"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      TabIndex        =   64
      Top             =   7080
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Add Object from File"
         Height          =   735
         Left            =   720
         TabIndex        =   67
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Create Torus File"
         Height          =   735
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Delete Active Object"
         Height          =   735
         Left            =   1440
         TabIndex        =   65
         Top             =   360
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1560
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Rendering"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   9960
      TabIndex        =   56
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command8 
         Caption         =   "DO PHONG"
         Height          =   735
         Left            =   2160
         TabIndex        =   73
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Frame Frame9 
         Caption         =   "Shading"
         Height          =   1695
         Left            =   120
         TabIndex        =   60
         Top             =   1440
         Width           =   1815
         Begin VB.OptionButton Option10 
            Caption         =   "Phong"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   1575
         End
         Begin VB.OptionButton Option9 
            Caption         =   "None"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Gouraud"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Flat"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Draw Mesh"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Draw Face Normals"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Draw Vertex Normals"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Translations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7560
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1695
         Begin VB.OptionButton Option6 
            Caption         =   "Translation"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Rotation"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Scaling"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Axis"
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1695
         Begin VB.OptionButton Option3 
            Caption         =   "Z"
            Height          =   255
            Left            =   1200
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Y"
            Height          =   255
            Left            =   600
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "X"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   375
         End
      End
      Begin VB.CommandButton Command20 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command19 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   16
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H8000000E&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   494
      ScaleMode       =   0  'User
      ScaleWidth      =   494
      TabIndex        =   14
      Top             =   120
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Camera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   9960
      TabIndex        =   13
      Top             =   3480
      Width           =   4215
      Begin VB.CommandButton Command29 
         Caption         =   "Delete Camera"
         Height          =   495
         Left            =   2160
         TabIndex        =   72
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Add Camera"
         Height          =   495
         Left            =   120
         TabIndex        =   71
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Next Camera"
         Height          =   495
         Left            =   2160
         TabIndex        =   70
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Previous Camera"
         Height          =   495
         Left            =   120
         TabIndex        =   69
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Frame Frame7 
         Caption         =   "Focus Point"
         Height          =   2535
         Left            =   2160
         TabIndex        =   42
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton Command16 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   48
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton Command17 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   47
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton Command25 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   46
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command24 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command23 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   44
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton Command22 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   43
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   54
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   50
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   49
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Camera Point"
         Height          =   2535
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton Command7 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton Command6 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   34
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton Command5 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton Command4 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   32
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            TabIndex        =   30
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   41
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   40
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   180
         Min             =   -180
         TabIndex        =   27
         Top             =   3240
         Width           =   3975
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Twist Angle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   55
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   3000
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Light"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton Command33 
         Caption         =   "Previous Light"
         Height          =   495
         Left            =   120
         TabIndex        =   77
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Next Light"
         Height          =   495
         Left            =   1200
         TabIndex        =   76
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Add Light"
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Delete Light"
         Height          =   495
         Left            =   1200
         TabIndex        =   74
         Top             =   3480
         Width           =   975
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1680
         Top             =   240
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ROTATE"
         Height          =   1695
         Left            =   1800
         TabIndex        =   26
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   2
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command15 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Rho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Theta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Phi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   2160
         Width           =   615
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public theWorld As New World
Public FrameNum As Long

Private Sub Check1_Click()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Check2_Click()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Check3_Click()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Command1_Click()
Dim newObj As New Object3D
Dim newcolor As New FaceColor

CommonDialog1.ShowOpen
If CommonDialog1.filename <> "" Then
    Set newObj = CreateObjectFromFile(CommonDialog1.filename, newcolor)
    Tesselate newObj
'    AddNewObject.Show 1
'    newObj.Color.r = val(AddNewObject.Text1.Text)
'    newObj.Color.g = val(AddNewObject.Text2.Text)
'    newObj.Color.b = val(AddNewObject.Text3.Text)
    theWorld.AddObject newObj
End If
MainForm.Label16(2).Caption = theWorld.ActiveObject.Kd
MainForm.Label16(3).Caption = theWorld.ActiveObject.Ks
MainForm.Label16(4).Caption = theWorld.ActiveObject.n
MainForm.Label16(5).Caption = theWorld.ActiveObject.hue
MainForm.Label16(6).Caption = theWorld.ActiveObject.Saturation
MainForm.Label16(7).Caption = theWorld.ActiveObject.Lightness
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Command10_Click()
theWorld.ActiveLight.AlterRho 50
Label10.Caption = theWorld.ActiveLight.Rho
End Sub

Private Sub Command11_Click()
theWorld.ActiveLight.AlterRho -50
Label10.Caption = theWorld.ActiveLight.Rho
End Sub

Private Sub Command12_Click()
theWorld.ActiveLight.Altertheta 5
Label11.Caption = theWorld.ActiveLight.Theta
End Sub

Private Sub Command13_Click()
theWorld.ActiveLight.Altertheta -5
Label11.Caption = theWorld.ActiveLight.Theta
End Sub

Private Sub Command14_Click()
theWorld.ActiveLight.AlterPhi 5
Label12.Caption = theWorld.ActiveLight.Phi
End Sub

Private Sub Command15_Click()
theWorld.ActiveLight.AlterPhi -5
Label12.Caption = theWorld.ActiveLight.Phi
End Sub

Private Sub Command16_Click()
theWorld.ActiveCamera.AlterFocusZ -50
Label13.Caption = theWorld.ActiveCamera.FocusPoint.Z
End Sub

Private Sub Command17_Click()
theWorld.ActiveCamera.AlterFocusZ 50
Label13.Caption = theWorld.ActiveCamera.FocusPoint.Z
End Sub

Private Sub Command18_Click()
dotorus 100, 50, 200
End Sub

Private Sub Command19_Click()
If Option6 Then
    If Option1 Then
        theWorld.ActiveObject.TranslateX 10
    ElseIf Option2 Then
        theWorld.ActiveObject.TranslateY 10
    ElseIf Option3 Then
        theWorld.ActiveObject.TranslateZ 10
    End If
ElseIf Option5 Then
    If Option1 Then
        theWorld.ActiveObject.rotateX 10
    ElseIf Option2 Then
        theWorld.ActiveObject.rotatey 10
    ElseIf Option3 Then
        theWorld.ActiveObject.rotateZ 10
    End If
ElseIf Option4 Then
    If Option1 Then
        theWorld.ActiveObject.ScaleX 0.1
        theWorld.ActiveObject.ScaleY 0.1
        theWorld.ActiveObject.ScaleZ 0.1
    ElseIf Option2 Then
        theWorld.ActiveObject.ScaleX 0.1
        theWorld.ActiveObject.ScaleY 0.1
        theWorld.ActiveObject.ScaleZ 0.1
    ElseIf Option3 Then
        theWorld.ActiveObject.ScaleX 0.1
        theWorld.ActiveObject.ScaleY 0.1
        theWorld.ActiveObject.ScaleZ 0.1
    End If
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Command2_Click()
theWorld.ActiveCamera.AlterCameraX 50
Label4.Caption = theWorld.ActiveCamera.CameraPoint.x
End Sub

Private Sub Command20_Click()
If Option6 Then
    If Option1 Then
        theWorld.ActiveObject.TranslateX -10
    ElseIf Option2 Then
        theWorld.ActiveObject.TranslateY -10
    ElseIf Option3 Then
        theWorld.ActiveObject.TranslateZ -10
    End If
ElseIf Option5 Then
    If Option1 Then
        theWorld.ActiveObject.rotateX -10
    ElseIf Option2 Then
        theWorld.ActiveObject.rotatey -10
    ElseIf Option3 Then
        theWorld.ActiveObject.rotateZ -10
    End If
ElseIf Option4 Then
    If Option1 Then
        theWorld.ActiveObject.ScaleX -0.1
        theWorld.ActiveObject.ScaleY -0.1
        theWorld.ActiveObject.ScaleZ -0.1
    ElseIf Option2 Then
        theWorld.ActiveObject.ScaleX -0.1
        theWorld.ActiveObject.ScaleY -0.1
        theWorld.ActiveObject.ScaleZ -0.1
    ElseIf Option3 Then
        theWorld.ActiveObject.ScaleX -0.1
        theWorld.ActiveObject.ScaleY -0.1
        theWorld.ActiveObject.ScaleZ -0.1
    End If
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Command21_Click()
If theWorld.NumberOfObjects > 0 Then
    theWorld.DeleteObject theWorld.ActiveObject.objectNum
End If
MainForm.Label16(2).Caption = theWorld.ActiveObject.Kd
MainForm.Label16(3).Caption = theWorld.ActiveObject.Ks
MainForm.Label16(4).Caption = theWorld.ActiveObject.n
MainForm.Label16(5).Caption = theWorld.ActiveObject.hue
MainForm.Label16(6).Caption = theWorld.ActiveObject.Saturation
MainForm.Label16(7).Caption = theWorld.ActiveObject.Lightness
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Command22_Click()
theWorld.ActiveCamera.AlterFocusY -50
Label14.Caption = theWorld.ActiveCamera.FocusPoint.y
End Sub

Private Sub Command23_Click()
theWorld.ActiveCamera.AlterFocusY 50
Label14.Caption = theWorld.ActiveCamera.FocusPoint.y
End Sub

Private Sub Command24_Click()
theWorld.ActiveCamera.AlterFocusX -50
Label17.Caption = theWorld.ActiveCamera.FocusPoint.x
End Sub

Private Sub Command25_Click()
theWorld.ActiveCamera.AlterFocusX 50
Label17.Caption = theWorld.ActiveCamera.FocusPoint.x
End Sub

Private Sub Command26_Click()
Dim curCam As Integer

curCam = theWorld.ActiveCamera.CameraNum
If curCam = 1 Then
    Set theWorld.ActiveCamera = theWorld.GetCamera(theWorld.NumberofCameras)
Else
    Set theWorld.ActiveCamera = theWorld.GetCamera(curCam - 1)
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
Label4.Caption = theWorld.ActiveCamera.CameraPoint.x
Label5.Caption = theWorld.ActiveCamera.CameraPoint.y
Label6.Caption = theWorld.ActiveCamera.CameraPoint.Z
Label17.Caption = theWorld.ActiveCamera.FocusPoint.x
Label14.Caption = theWorld.ActiveCamera.FocusPoint.y
Label13.Caption = theWorld.ActiveCamera.FocusPoint.Z
Label21.Caption = theWorld.ActiveCamera.TwistAngle
End Sub

Private Sub Command27_Click()
Dim curCam As Integer

curCam = theWorld.ActiveCamera.CameraNum
If curCam = theWorld.NumberofCameras Then
    Set theWorld.ActiveCamera = theWorld.GetCamera(1)
Else
    Set theWorld.ActiveCamera = theWorld.GetCamera(curCam + 1)
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
Label4.Caption = theWorld.ActiveCamera.CameraPoint.x
Label5.Caption = theWorld.ActiveCamera.CameraPoint.y
Label6.Caption = theWorld.ActiveCamera.CameraPoint.Z
Label17.Caption = theWorld.ActiveCamera.FocusPoint.x
Label14.Caption = theWorld.ActiveCamera.FocusPoint.y
Label13.Caption = theWorld.ActiveCamera.FocusPoint.Z
Label21.Caption = theWorld.ActiveCamera.TwistAngle
End Sub

Private Sub Command28_Click()
AddNewCamera.Show
AddNewCamera.SetFocus
End Sub

Private Sub Command29_Click()
If theWorld.NumberofCameras > 1 Then
    theWorld.DeleteCamera theWorld.ActiveCamera.CameraNum
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
Label4.Caption = theWorld.ActiveCamera.CameraPoint.x
Label5.Caption = theWorld.ActiveCamera.CameraPoint.y
Label6.Caption = theWorld.ActiveCamera.CameraPoint.Z
Label17.Caption = theWorld.ActiveCamera.FocusPoint.x
Label14.Caption = theWorld.ActiveCamera.FocusPoint.y
Label13.Caption = theWorld.ActiveCamera.FocusPoint.Z
Label21.Caption = theWorld.ActiveCamera.TwistAngle
End Sub

Private Sub Command3_Click()
theWorld.ActiveCamera.AlterCameraX -50
Label4.Caption = theWorld.ActiveCamera.CameraPoint.x
End Sub

Private Sub Command30_Click()
If theWorld.NumberofLights > 1 Then
    theWorld.DeleteLight theWorld.ActiveLight.LightNum
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
Label10.Caption = theWorld.ActiveLight.Rho
Label11.Caption = theWorld.ActiveLight.Theta
Label12.Caption = theWorld.ActiveLight.Phi
Label16(0).Caption = theWorld.ActiveLight.Ambient
Label16(1).Caption = theWorld.ActiveLight.Intensity
End Sub

Private Sub Command31_Click()
AddNewLightFrm.Show
AddNewLightFrm.SetFocus
End Sub

Private Sub Command32_Click()
Dim curLight As Integer

curLight = theWorld.ActiveLight.LightNum
If curLight = theWorld.NumberofLights Then
    Set theWorld.ActiveLight = theWorld.GetLight(1)
Else
    Set theWorld.ActiveLight = theWorld.GetLight(curLight + 1)
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
Label10.Caption = theWorld.ActiveLight.Rho
Label11.Caption = theWorld.ActiveLight.Theta
Label12.Caption = theWorld.ActiveLight.Phi
Label16(0).Caption = theWorld.ActiveLight.Ambient
Label16(1).Caption = theWorld.ActiveLight.Intensity
End Sub

Private Sub Command33_Click()
Dim curLight As Integer

curLight = theWorld.ActiveLight.LightNum
If curLight = 1 Then
    Set theWorld.ActiveLight = theWorld.GetLight(theWorld.NumberofLights)
Else
    Set theWorld.ActiveLight = theWorld.GetLight(curLight - 1)
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
Label10.Caption = theWorld.ActiveLight.Rho
Label11.Caption = theWorld.ActiveLight.Theta
Label12.Caption = theWorld.ActiveLight.Phi
Label16(0).Caption = theWorld.ActiveLight.Ambient
Label16(1).Caption = theWorld.ActiveLight.Intensity
End Sub

Private Sub Command34_Click()
Label23.Caption = theWorld.ActiveCamera.CameraPoint.x
Label24.Caption = theWorld.ActiveCamera.CameraPoint.y
Label25.Caption = theWorld.ActiveCamera.CameraPoint.Z
End Sub

Private Sub Command35_Click()
Label26.Caption = theWorld.ActiveCamera.CameraPoint.x
Label27.Caption = theWorld.ActiveCamera.CameraPoint.y
Label28.Caption = theWorld.ActiveCamera.CameraPoint.Z

End Sub

Private Sub Command36_Click()
'5 secs video = 150 frames (30 frames / sec)

Dim Dx As Double
Dim dy As Double
Dim dz As Double
Dim fso
Dim txtfile

FrameNum = 0
Dx = -(val(Label23.Caption) - val(Label26.Caption)) / 150
dy = -(val(Label24.Caption) - val(Label27.Caption)) / 150
dz = -(val(Label25.Caption) - val(Label28.Caption)) / 150
theWorld.ActiveCamera.CameraPoint.x = Label23.Caption
theWorld.ActiveCamera.CameraPoint.y = Label24.Caption
theWorld.ActiveCamera.CameraPoint.Z = Label25.Caption
'theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
'FrameNum = FrameNum + 1

theWorld.RasterDevice.Refresh
For i = 1 To 150
    theWorld.ActiveCamera.CameraPoint.x = theWorld.ActiveCamera.CameraPoint.x + Dx
    theWorld.ActiveCamera.CameraPoint.y = theWorld.ActiveCamera.CameraPoint.y + dy
    theWorld.ActiveCamera.CameraPoint.Z = theWorld.ActiveCamera.CameraPoint.Z + dz
    theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
    theWorld.RasterDevice.Refresh
    FrameNum = FrameNum + 1
    Command8_Click
Next

'Add pictures to Avi
RenderFrm.Picture1.Print "AVI Parameters..."
RenderFrm.Picture1.Refresh
Load RenderOpts
RenderOpts.cmdWriteAVI_Click
'Delete Pictures
RenderFrm.Picture1.Print "Deleting temp BMPs..."
RenderFrm.Picture1.Refresh
For i = 1 To 150
    Kill App.Path & "\test" & FixFiveDigit(CStr(i)) & ".bmp"
Next
RenderFrm.Picture1.Print "Finished!"
RenderFrm.Picture1.Refresh
End Sub

Private Sub Command4_Click()
theWorld.ActiveCamera.AlterCameraY 50
Label5.Caption = theWorld.ActiveCamera.CameraPoint.y
End Sub

Private Sub Command5_Click()
theWorld.ActiveCamera.AlterCameraY -50
Label5.Caption = theWorld.ActiveCamera.CameraPoint.y
End Sub

Private Sub Command6_Click()
theWorld.ActiveCamera.AlterCameraZ 50
Label6.Caption = theWorld.ActiveCamera.CameraPoint.Z
End Sub

Private Sub Command7_Click()
theWorld.ActiveCamera.AlterCameraZ -50
Label6.Caption = theWorld.ActiveCamera.CameraPoint.Z
End Sub

Private Sub Command8_Click()
RenderFrm.Show
RenderFrm.SetFocus

Set theWorld.RasterDevice = RenderFrm.Picture1
RenderFrm.Picture1.Print "Rendering...   " & CStr(FrameNum) & " of 150"
RenderFrm.Picture1.Refresh
theWorld.Raster False, False, False, False, False, True, False, False
Set theWorld.RasterDevice = MainForm.Picture1
'Save Pictures
SavePicture RenderFrm.Picture1.Image, App.Path & "\test" & FixFiveDigit(CStr(FrameNum)) & ".bmp"
End Sub

Private Sub Command9_Click()
If Command9.Caption = "ROTATE" Then
    Timer2.Enabled = True
    Command9.Caption = "STOP"
Else
    Timer2.Enabled = False
    Command9.Caption = "ROTATE"
End If
End Sub

Private Sub Form_Load()
Dim theLight As New Light
Dim theCamera As New Camera
Dim CPoint As New Vector3D
Dim FPoint As New Vector3D

StartCode

theWorld.Create Picture1

theLight.Create 500, 30, 60, 0.1, 0.8
theWorld.AddLight theLight

CPoint.x = 1000
CPoint.y = 800
CPoint.Z = 600
FPoint.x = 0
FPoint.y = 0
FPoint.Z = 0
theCamera.Create CPoint, FPoint, 0
theWorld.AddCamera theCamera

Label4.Caption = theWorld.ActiveCamera.CameraPoint.x
Label5.Caption = theWorld.ActiveCamera.CameraPoint.y
Label6.Caption = theWorld.ActiveCamera.CameraPoint.Z
Label17.Caption = theWorld.ActiveCamera.FocusPoint.x
Label14.Caption = theWorld.ActiveCamera.FocusPoint.y
Label13.Caption = theWorld.ActiveCamera.FocusPoint.Z
Label21.Caption = theWorld.ActiveCamera.TwistAngle
Label10.Caption = theWorld.ActiveLight.Rho
Label11.Caption = theWorld.ActiveLight.Theta
Label12.Caption = theWorld.ActiveLight.Phi
Label16(0).Caption = theWorld.ActiveLight.Ambient
Label16(1).Caption = theWorld.ActiveLight.Intensity
End Sub

Private Sub Form_Unload(Cancel As Integer)
For Each Form In Forms
    Unload Form
Next
End Sub

Private Sub HScroll1_Change()
theWorld.ActiveCamera.TwistAngle = HScroll1.Value
Label21.Caption = theWorld.ActiveCamera.TwistAngle
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub


Private Sub HScroll2_Change(Index As Integer)
If Index = 0 Then
    Label16(Index) = HScroll2(Index).Value / 1000
    theWorld.ActiveLight.Ambient = HScroll2(Index).Value / 1000
ElseIf Index = 1 Then
    Label16(Index) = HScroll2(Index).Value / 1000
    theWorld.ActiveLight.Intensity = HScroll2(Index).Value / 1000
ElseIf Index = 2 Then
    Label16(Index) = HScroll2(Index).Value / 1000
    theWorld.ActiveObject.Kd = HScroll2(Index).Value / 1000
ElseIf Index = 3 Then
    Label16(Index) = HScroll2(Index).Value / 1000
    theWorld.ActiveObject.Ks = HScroll2(Index).Value / 1000
ElseIf Index = 4 Then
    Label16(Index) = HScroll2(Index).Value
    theWorld.ActiveObject.n = HScroll2(Index).Value
ElseIf Index = 5 Then
    Label16(Index) = HScroll2(Index).Value
    theWorld.ActiveObject.hue = HScroll2(Index).Value
ElseIf Index = 6 Then
    Label16(Index) = HScroll2(Index).Value / 1000
    theWorld.ActiveObject.Saturation = HScroll2(Index).Value / 1000
ElseIf Index = 7 Then
    Label16(Index) = HScroll2(Index).Value / 1000
    theWorld.ActiveObject.Lightness = HScroll2(Index).Value / 1000
End If
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub HScroll2_Scroll(Index As Integer)
'HScroll2_Change Index
End Sub

Private Sub Label10_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Label11_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Label12_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Label13_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Label14_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Label16_Change(Index As Integer)
If Index = 0 Then
    HScroll2(Index).Value = val(Label16(Index).Caption * 1000)
'    theWorld.ActiveLight.Ambient = HScroll2(Index).Value
ElseIf Index = 1 Then
    HScroll2(Index).Value = val(Label16(Index).Caption * 1000)
'    theWorld.ActiveLight.Intensity = HScroll2(Index).Value
ElseIf Index = 2 Then
    HScroll2(Index).Value = val(Label16(Index).Caption * 1000)
'    theWorld.ActiveLight.Kd = HScroll2(Index).Value / 1000
ElseIf Index = 3 Then
    HScroll2(Index).Value = val(Label16(Index).Caption * 1000)
'    theWorld.ActiveLight.Ks = HScroll2(Index).Value / 1000
ElseIf Index = 4 Then
     HScroll2(Index).Value = val(Label16(Index).Caption)
'    theWorld.ActiveLight.n = HScroll2(Index).Value
ElseIf Index = 5 Then
     HScroll2(Index).Value = val(Label16(Index).Caption)
'    theWorld.ActiveLight.n = HScroll2(Index).Value
ElseIf Index = 6 Then
     HScroll2(Index).Value = val(Label16(Index).Caption * 1000)
'    theWorld.ActiveLight.n = HScroll2(Index).Value
ElseIf Index = 7 Then
     HScroll2(Index).Value = val(Label16(Index).Caption * 1000)
'    theWorld.ActiveLight.n = HScroll2(Index).Value
End If
End Sub

Private Sub Label17_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Label21_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
HScroll1.Value = val(Label21.Caption)
End Sub

Private Sub Label4_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True

End Sub

Private Sub Label5_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Label6_Change()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Option10_Click()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Option7_Click()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Option8_Click()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Option9_Click()
theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value, Option10.Value, True, True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
theWorld.ClickMe x, y
End Sub

Private Sub Text1_Change()
If (val(Text1.Text) > 0) And (val(Text1.Text) <= theWorld.NumberOfObjects) Then
    Set theWorld.ActiveObject = theWorld.GetObject(Text1.Text)
End If
End Sub

Private Sub Timer1_Timer()
'theWorld.ActiveCamera.Altertheta 3
'Label5.Caption = theWorld.ActiveCamera.Theta

'theWorld.ActiveCamera.AlterPhi 3
'Label6.Caption = theWorld.ActiveCamera.Phi

End Sub

Private Sub Timer2_Timer()
'theWorld.ActiveLight.AlterTheta 3
'Label11.Caption = theWorld.ActiveLight.Theta

theWorld.ActiveLight.AlterPhi -3
Label12.Caption = theWorld.ActiveLight.Phi

End Sub

Private Sub Timer3_Timer()
'theWorld.Raster Check1.Value, Check2.Value, Check3.Value, Option7.Value, Option8.Value,Option10.Value
End Sub
