VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "函数在线分析"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   0
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   12510
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12495
      Begin VB.Label Label1 
         Caption         =   "连接中......"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   200.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   12375
      End
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   9000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12500
      ExtentX         =   22049
      ExtentY         =   15875
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "连接中..."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   99.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3120
      TabIndex        =   1
      Top             =   3600
      Width           =   7575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
     Web.Navigate ("http://zh.numberempire.com/graphingcalculator.php")
DoEvents
While Web.Busy
    DoEvents
Wend
Web.Document.parentwindow.scrollby 0, 165
Web.Document.body.Scroll = "no"
Frame1.Visible = False
    

End Sub

