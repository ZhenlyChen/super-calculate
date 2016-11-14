VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Í¼±ê¿â"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   9075
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Image min 
      Height          =   480
      Index           =   1
      Left            =   3360
      Picture         =   "Form2.frx":08CA
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image min 
      Height          =   480
      Index           =   0
      Left            =   2880
      Picture         =   "Form2.frx":09B4
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image cmabout 
      Height          =   480
      Index           =   1
      Left            =   2400
      Picture         =   "Form2.frx":0A9E
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image cmabout 
      Height          =   480
      Index           =   0
      Left            =   1920
      Picture         =   "Form2.frx":0B3B
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image mini 
      Height          =   480
      Index           =   1
      Left            =   1440
      Picture         =   "Form2.frx":0BD8
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image mini 
      Height          =   480
      Index           =   0
      Left            =   960
      Picture         =   "Form2.frx":0C9C
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image closes 
      Height          =   480
      Index           =   1
      Left            =   480
      Picture         =   "Form2.frx":0D5A
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image closes 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":0E5D
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image new4 
      Height          =   720
      Index           =   1
      Left            =   7680
      Picture         =   "Form2.frx":0F55
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image new4 
      Height          =   720
      Index           =   0
      Left            =   7680
      Picture         =   "Form2.frx":112F
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image new3 
      Height          =   720
      Index           =   1
      Left            =   6960
      Picture         =   "Form2.frx":12DA
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image new3 
      Height          =   720
      Index           =   0
      Left            =   6960
      Picture         =   "Form2.frx":14DB
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image new2 
      Height          =   720
      Index           =   1
      Left            =   7680
      Picture         =   "Form2.frx":16AD
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image new2 
      Height          =   720
      Index           =   0
      Left            =   7680
      Picture         =   "Form2.frx":18B8
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image new1 
      Height          =   720
      Index           =   1
      Left            =   6960
      Picture         =   "Form2.frx":1A1C
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image new1 
      Height          =   720
      Index           =   0
      Left            =   6960
      Picture         =   "Form2.frx":1CCF
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image cmcp 
      Height          =   720
      Index           =   2
      Left            =   6240
      Picture         =   "Form2.frx":1FAF
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image cmcp 
      Height          =   720
      Index           =   1
      Left            =   6240
      Picture         =   "Form2.frx":21A3
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image cmcp 
      Height          =   720
      Index           =   0
      Left            =   6240
      Picture         =   "Form2.frx":23BF
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image cmecp 
      Height          =   480
      Index           =   2
      Left            =   5760
      Picture         =   "Form2.frx":25A9
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image cmecp 
      Height          =   480
      Index           =   1
      Left            =   5760
      Picture         =   "Form2.frx":2734
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image cmecp 
      Height          =   480
      Index           =   0
      Left            =   5760
      Picture         =   "Form2.frx":28CC
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image cmzf 
      Height          =   720
      Index           =   2
      Left            =   5040
      Picture         =   "Form2.frx":2A54
      ToolTipText     =   "´¢´æÊý×Ö"
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image cmzf 
      Height          =   720
      Index           =   1
      Left            =   5040
      Picture         =   "Form2.frx":2B2F
      ToolTipText     =   "´¢´æÊý×Ö"
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image cmzf 
      Height          =   720
      Index           =   0
      Left            =   5040
      Picture         =   "Form2.frx":2C63
      ToolTipText     =   "´¢´æÊý×Ö"
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image cmsave 
      Height          =   720
      Index           =   2
      Left            =   4320
      Picture         =   "Form2.frx":2D3E
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image cmsave 
      Height          =   720
      Index           =   1
      Left            =   4320
      Picture         =   "Form2.frx":2E52
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image cmsave 
      Height          =   720
      Index           =   0
      Left            =   4320
      Picture         =   "Form2.frx":2F86
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image cmpoint 
      Height          =   720
      Index           =   2
      Left            =   3600
      Picture         =   "Form2.frx":309C
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image cmpoint 
      Height          =   720
      Index           =   1
      Left            =   3600
      Picture         =   "Form2.frx":3128
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image cmpoint 
      Height          =   720
      Index           =   0
      Left            =   3600
      Picture         =   "Form2.frx":31BF
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image enter 
      Height          =   720
      Index           =   2
      Left            =   2880
      Picture         =   "Form2.frx":324B
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image enter 
      Height          =   720
      Index           =   1
      Left            =   2880
      Picture         =   "Form2.frx":32DC
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image enter 
      Height          =   720
      Index           =   0
      Left            =   2880
      Picture         =   "Form2.frx":338C
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image sign2 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "Form2.frx":341D
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image sign1 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "Form2.frx":34F0
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image sign0 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "Form2.frx":35B3
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image sign2 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "Form2.frx":3686
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image sign1 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "Form2.frx":37DB
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image sign0 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "Form2.frx":3906
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image sign2 
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "Form2.frx":3A4C
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image sign1 
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "Form2.frx":3AD8
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image sign0 
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "Form2.frx":3B6B
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image sign2 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":3BF7
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image sign1 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":3CB7
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image sign0 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":3D7D
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image cmleft 
      Height          =   720
      Index           =   2
      Left            =   7920
      Picture         =   "Form2.frx":3E3D
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image cmleft 
      Height          =   720
      Index           =   1
      Left            =   7920
      Picture         =   "Form2.frx":3F1A
      Top             =   720
      Width           =   720
   End
   Begin VB.Image cmleft 
      Height          =   720
      Index           =   0
      Left            =   7920
      Picture         =   "Form2.frx":4003
      Top             =   0
      Width           =   720
   End
   Begin VB.Image cmcls 
      Height          =   720
      Index           =   2
      Left            =   7200
      Picture         =   "Form2.frx":40E0
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image cmcls 
      Height          =   720
      Index           =   1
      Left            =   7200
      Picture         =   "Form2.frx":42B0
      Top             =   720
      Width           =   720
   End
   Begin VB.Image cmcls 
      Height          =   720
      Index           =   0
      Left            =   7200
      Picture         =   "Form2.frx":44AC
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   9
      Left            =   6480
      Picture         =   "Form2.frx":4670
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   8
      Left            =   5760
      Picture         =   "Form2.frx":484C
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   7
      Left            =   5040
      Picture         =   "Form2.frx":4A30
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   6
      Left            =   4320
      Picture         =   "Form2.frx":4BBE
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   5
      Left            =   3600
      Picture         =   "Form2.frx":4D9F
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   4
      Left            =   2880
      Picture         =   "Form2.frx":4F6D
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "Form2.frx":5118
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "Form2.frx":52E5
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "Form2.frx":5490
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":559E
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   9
      Left            =   6480
      Picture         =   "Form2.frx":576B
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   8
      Left            =   5760
      Picture         =   "Form2.frx":5971
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   7
      Left            =   5040
      Picture         =   "Form2.frx":5B75
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   6
      Left            =   4320
      Picture         =   "Form2.frx":5D15
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   5
      Left            =   3600
      Picture         =   "Form2.frx":5F1E
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   4
      Left            =   2880
      Picture         =   "Form2.frx":6109
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "Form2.frx":62CF
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "Form2.frx":64B8
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "Form2.frx":6692
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":67AB
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   9
      Left            =   6480
      Picture         =   "Form2.frx":69A1
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   8
      Left            =   5760
      Picture         =   "Form2.frx":6B6F
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   7
      Left            =   5040
      Picture         =   "Form2.frx":6D50
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   6
      Left            =   4320
      Picture         =   "Form2.frx":6EDB
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   5
      Left            =   3600
      Picture         =   "Form2.frx":70B6
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   4
      Left            =   2880
      Picture         =   "Form2.frx":727B
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "Form2.frx":741F
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "Form2.frx":75E3
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "Form2.frx":778E
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "Form2.frx":789A
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
