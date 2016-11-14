VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00353535&
   BorderStyle     =   0  'None
   Caption         =   "¼ÆËã¼ÇÂ¼"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   6165
   ScaleWidth      =   7215
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   120
      Picture         =   "Form6.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Tip 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   4815
   End
   Begin VB.Label up 
      BackStyle       =   0  'Transparent
      Caption         =   "±£´æ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image closes 
      Height          =   480
      Left            =   6600
      Picture         =   "Form6.frx":1194
      ToolTipText     =   "¹Ø±Õ"
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "¼ÆËã¼ÇÂ¼"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   80
      Width           =   1575
   End
   Begin VB.Image min 
      Height          =   480
      Left            =   6000
      Picture         =   "Form6.frx":128C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Çå¿Õ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub closes_Click()
Unload Me
End Sub

Private Sub closes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closes = Form2.closes(1)
min = Form2.min(0)
End Sub

Private Sub Label2_Click()
Text1 = ""
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF00&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFF&
up.ForeColor = &HFFFFFF
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFFFF
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closes = Form2.closes(0)
min = Form2.min(0)
Label2.ForeColor = &HFFFFFF
up.ForeColor = &HFFFFFF
End Sub

Private Sub Tip_Click()
Shell "explorer " & App.Path, 1
End Sub

Private Sub up_Click()
Dim times As String
times = Hour(Time) & Minute(Time) & Second(Time)
Open App.Path & "\history" & times & ".txt" For Output As #1
Print #1, Text1
Close
Tip = "ÒÑ±£´æ£¬µã»÷´Ë´¦´ò¿ªÎÄ¼þ¼Ð"
Tip.ToolTipText = App.Path & "\history" & times & ".txt"
End Sub

Private Sub up_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
up.ForeColor = &HFF00&
End Sub

Private Sub up_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
up.ForeColor = &HFFFF&
Label2.ForeColor = &HFFFFFF
End Sub

Private Sub up_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
up.ForeColor = &HFFFFFF
End Sub
Private Sub min_Click()
Me.WindowState = 1
End Sub

Private Sub min_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closes = Form2.closes(0)
min = Form2.min(1)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closes = Form2.closes(0)
min = Form2.min(0)
Label2.ForeColor = &HFFFFFF
up.ForeColor = &HFFFFFF
End Sub

