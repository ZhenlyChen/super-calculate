VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00353535&
   BorderStyle     =   0  'None
   Caption         =   "智能分析"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4095
   ScaleWidth      =   7260
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5160
      Top             =   240
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "复制结果"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6240
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "粘贴"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   120
      Picture         =   "Form3.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Image new3 
      Height          =   720
      Left            =   6240
      Picture         =   "Form3.frx":1194
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "请在下面输入一条算式，系统将智能分析得出结果                      可用字符：+-*/×÷01234567890 ( ) { } [ ]【不然程序可能崩溃】"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "智能分析"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   150
      Width           =   1815
   End
   Begin VB.Image closes 
      Height          =   480
      Left            =   6600
      Picture         =   "Form3.frx":1366
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image min 
      Height          =   480
      Left            =   6000
      Picture         =   "Form3.frx":145E
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As String
Private Sub closes_Click()
Unload Me
End Sub

Private Sub Label2_Click()

Text1.Text = Clipboard.GetText
End Sub
Private Sub Label3_Click()
Clipboard.Clear
Clipboard.SetText Text1 & "=" & Text2
Label5 = "已复制到粘贴板"
Timer1.Enabled = True
End Sub

Private Sub min_Click()
Me.WindowState = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub


Public Function JC(ByVal mSS As String) As String
On Err GoTo CuoWu
    JC = Replace(Replace(mSS, " ", ""), "　", "")
    JC = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(JC, "}", ")"), "{", "("), "]", ")"), "[", "("), "=", ""), "。", "."), "、", "/"), "\", "/"), "X", "*"), "x", "*"), "（", "("), "）", ")"), "×", "*"), "÷", "/"), "＋", "+"), "－", "-"), "＊", "*"), "／", "/")
    JC = Replace(JC, "０", "0")
    JC = Replace(JC, "１", "1")
    JC = Replace(JC, "２", "2")
    JC = Replace(JC, "３", "3")
    JC = Replace(JC, "４", "4")
    JC = Replace(JC, "５", "5")
    JC = Replace(JC, "６", "6")
    JC = Replace(JC, "７", "7")
    JC = Replace(JC, "８", "8")
    JC = Replace(JC, "９", "9")

Exit Function
CuoWu:
Text2 = "分析不了啦！"
End Function





Public Function JKH(ByVal mSS As String) As String
On Err GoTo CuoWu
    Dim iss(2) As String
    Dim iTemp As Variant
    If InStr(mSS, "(") <> 0 Then
        iss(2) = Mid(mSS, InStr(mSS, ")") + 1)
        iss(0) = Left(mSS, Len(mSS) - Len(iss(2)) - 1)
        iTemp = Split(iss(0), "(")
        iss(1) = iTemp(UBound(iTemp))
        iss(0) = Left(iss(0), Len(iss(0)) - Len(iss(1)) - 1)
        JKH = JKH(iss(0) & SJJ(SCC(iss(1))) & iss(2))
    Else
        JKH = SJJ(SCC(mSS))
    End If
    Exit Function
CuoWu:
Text2 = "分析不了啦！"
End Function

Public Function SCC(ByVal mSS As String) As String
On Err GoTo CuoWu
    Dim iss(2) As String
    Dim iTemp As Variant
    Dim itemp0 As Variant
    Dim itemp1 As Variant
    Dim Sum1 As Double, Sum2 As Double, FH As String
    Dim sum0 As Double
    If InStr(mSS, "*") = 0 And InStr(mSS, "/") = 0 Then
        SCC = mSS
        Exit Function
    End If
    iTemp = Split(Replace(mSS, "/", "*"), "*")
    itemp0 = Split(Replace(iTemp(0), "-", "+"), "+")
    itemp1 = Split(Replace(iTemp(1), "-", "+"), "+")
    Sum1 = itemp0(UBound(itemp0))
    Sum2 = itemp1(0)
    iss(0) = Left(iTemp(0), Len(iTemp(0)) - Len(CStr(Sum1)))
    iss(2) = Mid(mSS, Len(iTemp(0)) + Len(CStr(Sum2)) + 2)
    FH = Mid(mSS, Len(iTemp(0)) + 1, 1)
    
    Select Case FH
    Case "*": sum0 = Sum1 * Sum2
    Case "/": sum0 = Sum1 / Sum2
    End Select
    SCC = SCC(iss(0) & CStr(sum0) & iss(2))
    Exit Function
CuoWu:
Text2 = "分析不了啦！"
End Function

Public Function SJJ(ByVal mSS As String) As String
On Err GoTo CuoWu
    Dim iss(2) As String
    Dim iTemp As Variant
    Dim itemp0 As Variant
    Dim itemp1 As Variant
    Dim Sum1 As Double, Sum2 As Double, FH As String
    Dim sum0 As Double
    If InStr(mSS, "+") = 0 And InStr(mSS, "-") <= 1 Then
        SJJ = mSS
        Exit Function
    End If
    
    mSS = Replace(Replace(Replace(Replace(mSS, "++", "+"), "+-", "-"), "-+", "-"), "--", "+")
    iTemp = Split(Replace(mSS, "-", "+"), "+")
    If IsNumeric(iTemp(0)) = False Then GoTo CuoWu
    If IsNumeric(iTemp(1)) = False Then GoTo CuoWu
    Sum1 = iTemp(0)
    Sum2 = iTemp(1)
    iss(2) = Mid(mSS, Len(iTemp(0)) + Len(iTemp(1)) + 2)
    FH = Mid(mSS, Len(iTemp(0)) + 1, 1)
    
    Select Case FH
    Case "+": sum0 = Sum1 + Sum2
    Case "-": sum0 = Sum1 - Sum2
    End Select
    SJJ = SJJ(iss(0) & CStr(sum0) & iss(2))
    Exit Function
CuoWu:
Text2 = "分析不了啦！"
End Function

Private Sub new3_Click()
Text2 = JKH(JC(Text1))
If Text2 = "" Then Text2 = "这是神马东西啊"
End Sub
'=====================================界面动画效果
Private Sub ccls()
closes = Form2.closes(0)
min = Form2.min(0)
new3 = Form2.new3(0)
Label2.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
End Sub
Private Sub closes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
closes = Form2.closes(1)
End Sub




Private Sub new3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
new3 = Form2.new3(1)
End Sub
Private Sub min_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
min = Form2.min(1)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF00&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
Label2.ForeColor = &HFFFF&
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFFFF
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF00&
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
Label3.ForeColor = &HFFFF&
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFFFF
End Sub
Private Sub Text1_Change()
If InStr(Text1, "=") > 0 Then
Text1 = Left(Text1, InStr(Text1, "=") - 1)
End If
End Sub

Private Sub Timer1_Timer()
Label5 = ""
Timer1.Enabled = False
End Sub
