VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super计算"
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6600
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   3000
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   400
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6720
      Top             =   6240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "查看记录"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "精确小数"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6360
      Width           =   975
   End
   Begin VB.Image cmabout 
      Height          =   480
      Left            =   4800
      ToolTipText     =   "关于"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image mini 
      Height          =   480
      Left            =   5400
      ToolTipText     =   "最小化到托盘"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image closes 
      Height          =   480
      Left            =   6000
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   105
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Super计算器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image news4 
      Height          =   720
      Left            =   4440
      Picture         =   "Form1.frx":1194
      Top             =   5520
      Width           =   2025
   End
   Begin VB.Image news3 
      Height          =   720
      Left            =   4440
      Picture         =   "Form1.frx":17AB
      Top             =   4680
      Width           =   2025
   End
   Begin VB.Image new4 
      Height          =   720
      Left            =   3720
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image new3 
      Height          =   720
      Left            =   3720
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image news2 
      Height          =   720
      Left            =   4440
      Picture         =   "Form1.frx":1DF4
      Top             =   3840
      Width           =   2025
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   4560
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image new2 
      Height          =   720
      Left            =   3720
      Top             =   3840
      Width           =   720
   End
   Begin VB.Image news1 
      Height          =   720
      Left            =   4440
      Picture         =   "Form1.frx":2470
      Top             =   3000
      Width           =   2025
   End
   Begin VB.Image new1 
      Height          =   720
      Left            =   3720
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label Tip 
      BackStyle       =   0  'Transparent
      Caption         =   "鼠标悬停在按钮上可看提示"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   6495
   End
   Begin VB.Image cmcp 
      Height          =   720
      Left            =   5760
      ToolTipText     =   "复制"
      Top             =   1320
      Width           =   720
   End
   Begin VB.Image cmecp 
      Height          =   480
      Left            =   6000
      ToolTipText     =   "复制(查看全部)"
      Top             =   720
      Width           =   480
   End
   Begin VB.Image cmzf 
      Height          =   735
      Left            =   1800
      ToolTipText     =   "改变正负号（Q）"
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label save 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   5280
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label save 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   4005
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label save 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   5280
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "储存"
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label save 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   4005
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image cmsave 
      Height          =   735
      Left            =   960
      ToolTipText     =   "储存数字(S)"
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label EELED 
      BackColor       =   &H00000000&
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
      Height          =   480
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "计算过程"
      Top             =   720
      Width           =   6345
   End
   Begin VB.Label ELED 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   6840
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image cmpoint 
      Height          =   720
      Left            =   960
      ToolTipText     =   "小数点"
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image enter 
      Height          =   720
      Left            =   2640
      ToolTipText     =   "计算结果"
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image sign 
      Height          =   720
      Index           =   3
      Left            =   2640
      ToolTipText     =   "除"
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image sign 
      Height          =   720
      Index           =   2
      Left            =   2640
      ToolTipText     =   "乘"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image sign 
      Height          =   720
      Index           =   1
      Left            =   2640
      ToolTipText     =   "减"
      Top             =   3840
      Width           =   720
   End
   Begin VB.Image sign 
      Height          =   720
      Index           =   0
      Left            =   2640
      ToolTipText     =   "加"
      Top             =   4680
      Width           =   720
   End
   Begin VB.Label HLED 
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   7320
      Width           =   2535
   End
   Begin VB.Label LED 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "显示屏"
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Image cmleft 
      Height          =   735
      Left            =   1800
      ToolTipText     =   "退格"
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image cmcls 
      Height          =   720
      Left            =   120
      ToolTipText     =   "全部清空(Delete)"
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   9
      Left            =   1800
      ToolTipText     =   "9"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   8
      Left            =   960
      ToolTipText     =   "8"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   7
      Left            =   120
      ToolTipText     =   "7"
      Top             =   3000
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   6
      Left            =   1800
      ToolTipText     =   "6"
      Top             =   3840
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   5
      Left            =   960
      ToolTipText     =   "5"
      Top             =   3840
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   4
      Left            =   120
      ToolTipText     =   "4"
      Top             =   3840
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   3
      Left            =   1800
      ToolTipText     =   "3"
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   2
      Left            =   960
      ToolTipText     =   "2"
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   1
      Left            =   120
      ToolTipText     =   "1"
      Top             =   4680
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   0
      Left            =   120
      ToolTipText     =   "0"
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Top             =   720
      Width           =   6375
   End
   Begin VB.Menu F00 
      Caption         =   "托盘"
      Visible         =   0   'False
      Begin VB.Menu F01 
         Caption         =   "显示主窗体"
      End
      Begin VB.Menu F02 
         Caption         =   "退出本程序"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
 Dim strup, strcls, signs  As String
Dim numone As Single
Dim nsave(0 To 3) As Single
Dim numled, numsave As Integer
Dim WindowTop, WindowLeft
Dim OldSign As String
Function WindowStyle()
With nfIconData
.hWnd = Me.hWnd
.uID = Me.Icon
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon.Handle
.szTip = "Super计算" & vbNullChar
.cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)
Me.Hide
End Function

Private Sub cmabout_Click()
Form4.Show
Form4.Left = Me.Left
Form4.Top = Me.Top + Me.Height - Form4.Height
End Sub



Private Sub Form_Resize()
WindowTop = Me.Top
WindowLeft = Me.Left
If Me.WindowState = 1 Then
WindowStyle
End If
End Sub


'======================================================================================================窗体事件
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
Case WM_LBUTTONDBLCLK
ShowWindow Me.hWnd, SW_RESTORE
Me.Top = WindowTop
Me.Left = WindowLeft
Me.SetFocus
Case WM_RBUTTONUP
PopupMenu F00
End Select
ccls
End Sub
Private Sub F01_Click()
ShowWindow Me.hWnd, SW_RESTORE
Me.Top = WindowTop
Me.Left = WindowLeft
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub

Private Sub F02_Click()
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End
End Sub
Private Sub Form_Load()

Dim H As Long
H = GetWindowLong(Me.hWnd, GWL_STYLE)
SetWindowLong Me.hWnd, GWL_STYLE, H And Not WS_CAPTION
Me.Refresh
    strup = 0
    strcls = 0
    ccls
    numsave = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
Text1.Visible = False
EELED.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
'======================================================================================================键盘触发事件

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Text1.Visible = False
EELED.Visible = True
ccls
    If KeyCode = 229 Then MsgBox ("请关闭输入法或切换为英文状态再输入")
If 95 < Int(KeyCode) And Int(KeyCode) < 106 Then
    num_Click (Int(KeyCode) - 96)
            For nu = 0 To 9
                If nu = Int(Int(KeyCode) - 96) Then
                  num(nu) = Form2.Image3(nu)
                End If
            Next
    strup = 0
End If
If 47 < Int(KeyCode) And Int(KeyCode) < 58 Then
    num_Click (Int(KeyCode) - 48)
            For nu = 0 To 9
                If nu = Int(Int(KeyCode) - 48) Then
                  num(nu) = Form2.Image3(nu)
                End If
            Next
    strup = 0
End If

If KeyCode = 8 Then
cmleft_Click
cmleft = Form2.cmleft(2)
End If

If KeyCode = 190 Or KeyCode = 110 Then
cmpoint_Click
cmpoint = Form2.cmpoint(2)
End If

If KeyCode = 107 Then
sign_Click (0)
sign(0) = Form2.sign2(0)
End If

If KeyCode = 109 Or KeyCode = 189 Then
sign_Click (1)
sign(1) = Form2.sign2(1)
End If

If KeyCode = 81 Then
cmzf_Click
cmzf = Form2.cmzf(2)
End If

If KeyCode = 106 Or KeyCode = 88 Then
sign_Click (2)
sign(2) = Form2.sign2(2)
End If

If KeyCode = 111 Or KeyCode = 191 Then
sign_Click (3)
sign(3) = Form2.sign2(3)
End If

If KeyCode = 13 Then
enter_Click
enter = Form2.enter(2)
End If

If KeyCode = 46 Then
cmcls_Click
cmcls = Form2.cmcls(2)
End If
If KeyCode = 83 Then
cmsave_Click
cmsave = Form2.cmsave(2)
End If
If KeyCode = 187 Then
enter_Click
enter = Form2.enter(2)
    If Shift = 1 Then
    sign_Click (0)
    sign(0) = Form2.sign2(0)
    End If
End If
End Sub

'======================================================================================================其他触发事件
Private Sub HLED_Change()
LED = Right(HLED, 12)
Text1.Visible = False
EELED.Visible = True
End Sub
Private Sub ELED_Change()
EELED = Right(ELED, 27)
numled = Len(ELED)
Text1 = ELED
Text1.Visible = False
EELED.Visible = True
End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub Label3_Click()
Form6.Show

End Sub

Private Sub Label4_Click()
If LED <> "" Then
nums = InputBox("请输入精确的小数位【0~15】")
    If nums <> "" Then
    If nums = "0" Then LED = Format(LED, "0")
    If nums = "1" Then LED = Format(LED, "0.0")
    If nums = "2" Then LED = Format(LED, "0.00")
    If nums = "3" Then LED = Format(LED, "0.000")
    If nums = "4" Then LED = Format(LED, "0.0000")
    If nums = "5" Then LED = Format(LED, "0.00000")
    If nums = "6" Then LED = Format(LED, "0.000000")
    If nums = "7" Then LED = Format(LED, "0.0000000")
    If nums = "8" Then LED = Format(LED, "0.00000000")
    If nums = "9" Then LED = Format(LED, "0.000000000")
    If nums = "10" Then LED = Format(LED, "0.0000000000")
    If nums = "11" Then LED = Format(LED, "0.00000000000")
    If nums = "12" Then LED = Format(LED, "0.000000000000")
    If nums = "13" Then LED = Format(LED, "0.0000000000000")
    If nums = "14" Then LED = Format(LED, "0.00000000000000")
    If nums = "15" Then LED = Format(LED, "0.000000000000000")
    End If
End If
End Sub

Private Sub new3_Click()
Form3.Show
End Sub

Private Sub news1_Click()
MsgBox ("研发中......【“关于 — 检查更新”升级最新版本获得新功能】")
End Sub
Private Sub new1_Click()
MsgBox ("研发中......【“关于 — 检查更新”升级最新版本获得新功能】")
End Sub
Private Sub news2_Click()
Form7.Show
End Sub
Private Sub new2_Click()
Form7.Show
End Sub
Private Sub news3_Click()
Form3.Show
End Sub
Private Sub new4_Click()
Form5.Show
End Sub
Private Sub news4_Click()
Form5.Show
End Sub

Private Sub Timer1_Timer()
Tip = "鼠标悬停在按钮上可看提示"
Timer1 = False
End Sub

'=====================================================================================================按钮触发事件

Private Sub mini_Click()
Me.WindowState = 1
End Sub
Private Sub closes_Click()
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End
End Sub

Private Sub cmecp_Click()


Text1.Visible = True
EELED.Visible = False
Clipboard.Clear
Clipboard.SetText ELED
Tip = "已复制到粘贴板上了哦！\(^o^)/~"
Timer1 = True
End Sub
Private Sub cmcp_Click()
Clipboard.Clear
Clipboard.SetText LED
Tip = "已复制到粘贴板上了哦！\(^o^)/~"
Timer1 = True
End Sub
Private Sub cmzf_Click()
 If LED <> "+" And LED <> "-" And LED <> "×" And LED <> "÷" And LED <> "" Then
    If strcls = 1 Then
    ELED = ""
    numone = "0"
    strcls = 0
OldSign = ""
    End If

    If InStrRev(HLED, ".") > 0 Then HLED = Format(HLED, "0." & String(Len(HLED) - InStrRev(HLED, "."), "0"))
        If LED <> "+" And LED <> "-" And LED <> "×" And LED <> "÷" And LED <> "" And ELED <> "" And LED <> ELED Then
                ELED = Left(ELED, Len(ELED) - Len(LED))
                If CSng(LED) > 0 Then
                    HLED = -CSng(HLED)
                Else
                    HLED = Abs(CSng(HLED))
                End If
                ELED = ELED & LED
        Else
                ELED = ELED
                If CSng(LED) > 0 Then
                    HLED = -CSng(HLED)
                Else
                    HLED = Abs(CSng(HLED))
                End If
                ELED = elde & LED
        End If
        

 End If
End Sub
Private Sub cmsave_Click()
If LED <> "" Then
    If Right(LED, 1) = "+" Or Right(LED, 1) = "-" Or Right(LED, 1) = "×" Or Right(LED, 1) = "÷" Then
    Else
        If numsave = 4 Then numsave = 0

        
        If Len(LED) > 7 Then
            save(numsave) = Left(LED, 7) & "..."
        Else
            save(numsave) = Left(LED, 7)
        End If
        nsave(numsave) = CSng(LED)
        save(numsave).ToolTipText = nsave(numsave)

        numsave = numsave + 1
        Tip = "已将数字储存在" & numsave & "号储存区里，点击储存区可以读取啦！"
        Timer1 = True
    End If
Else
    Tip = "显示区里面没有数字哦！"
    Timer1 = True
End If
End Sub
Private Sub enter_Click()
ccls
If HLED = "." Then HLED = "0"
If HLED <> "" Then
If strcls = 0 Then
ELED = ELED + "="
            Select Case signs
                Case 0
                    HLED = CStr(CSng(numone) + CSng(HLED))
                Case 1
                   HLED = CStr(CSng(numone) - CSng(HLED))
                Case 2
                   HLED = CStr((CSng(numone) * CSng(HLED)))
                Case 3
                If CSng(HLED) <> 0 Then
                HLED = CStr(CSng(numone) / CSng(HLED))
                End If
            End Select
            ELED = ELED + LED
    strcls = 1
    Form6.Text1 = Form6.Text1 + ELED + vbCrLf
End If
End If
End Sub



Private Sub save1_Click()

End Sub

Private Sub save_Click(Index As Integer)
If save(Index) <> "" Then

    If strcls = 1 Then
        HLED = ""
        ELED = ""
        numone = "0"
        strcls = 0
OldSign = ""
    End If
    If LED <> "＋" And LED <> "-" And LED <> "×" And LED <> "÷" Then
       ELED = Left(ELED, Len(ELED) - Len(LED)) & nsave(Index)
    Else
       ELED = ELED & nsave(Index)
    End If

    HLED = nsave(Index)
End If
End Sub

Private Sub sign_Click(Index As Integer)
If Right(ELED, 1) <> "=" Then
If Len(HLED) < 39 Then
If strcls = 1 Then
numone = HLED
ELED = HLED
End If
If HLED = "." Then HLED = "0"
'If signs <> "" Then signs = CStr(Index)
        If CStr(numone) = "0" Then
            If HLED = "" Then
            numone = "0"
            ELED = "0"
            Else
            numone = HLED
            End If
            HLED = ""
            Else
             If strcls = 0 And HLED <> "" Then
                  Select Case signs
                     Case 0
                         numone = numone + CSng(HLED)
                     Case 1
                        numone = numone - CSng(HLED)
                     Case 2
                        numone = numone * CSng(HLED)
                     Case 3
                      
                        If CSng(HLED) <> 0 Then
                       numone = numone / CSng(HLED)
                End If
                  End Select
              
            End If
            
        End If
 HLED = ""
If Right(ELED, 1) = "+" Or Right(ELED, 1) = "-" Or Right(ELED, 1) = "×" Or Right(ELED, 1) = "÷" Then ELED = Left(ELED, Len(ELED) - 1)

    If Index = 0 Then ELED = ELED + "+"
    If Index = 1 Then ELED = ELED + "-"
    If Index = 2 Then ELED = ELED + "×"
    If Index = 3 Then ELED = ELED + "÷"
    If Index = 0 Then LED = "+"
    If Index = 1 Then LED = "-"
    If Index = 2 Then LED = "×"
    If Index = 3 Then LED = "÷"

    If strcls = 1 Then strcls = 0
signs = CStr(Index)
Else
Tip = "你输入这么长存心捣乱的吧！重置！！哈哈哈！"
cmcls_Click
End If
Else
cmcls_Click
Tip = "格式错误，已重置，哇哈哈嘻嘻！"
End If

End Sub
Private Sub num_Click(Index As Integer)
If Right(ELED, 1) = "+" Then OldSign = OldSign & "1"
If Right(ELED, 1) = "-" Then OldSign = OldSign & "2"
If Right(ELED, 1) = "×" Then OldSign = OldSign & "3"
If Right(ELED, 1) = "÷" Then OldSign = OldSign & "4"

If InStr(OldSign, "13") > 0 Then
    OldSign = Replace(OldSign, "13", "1,3")
    ELED = Left(ELED, Len(ELED) - 1)
    If Left(ELED, 1) = "[" Then
        ELED = "{" & ELED & "}" & "×"
    Else
        If Left(ELED, 1) = "(" Then
            ELED = "[" & ELED & "]" & "×"
        Else
            ELED = "(" & ELED & ")" & "×"
        End If
    End If
End If

If InStr(OldSign, "23") > 0 Then
    OldSign = Replace(OldSign, "23", "2,3")
    ELED = Left(ELED, Len(ELED) - 1)
    If Left(ELED, 1) = "[" Then
        ELED = "{" & ELED & "}" & "×"
    Else
        If Left(ELED, 1) = "(" Then
            ELED = "[" & ELED & "]" & "×"
        Else
            ELED = "(" & ELED & ")" & "×"
        End If
    End If
End If

If InStr(OldSign, "24") > 0 Then
    OldSign = Replace(OldSign, "24", "2,4")
    ELED = Left(ELED, Len(ELED) - 1)
    If Left(ELED, 1) = "[" Then
        ELED = "{" & ELED & "}" & "÷"
    Else
        If Left(ELED, 1) = "(" Then
            ELED = "[" & ELED & "]" & "÷"
        Else
            ELED = "(" & ELED & ")" & "÷"
        End If
    End If
End If

If InStr(OldSign, "14") > 0 Then
    OldSign = Replace(OldSign, "14", "1,4")
    ELED = Left(ELED, Len(ELED) - 1)
    If Left(ELED, 1) = "[" Then
        ELED = "{" & ELED & "}" & "÷"
    Else
        If Left(ELED, 1) = "(" Then
            ELED = "[" & ELED & "]" & "÷"
        Else
            ELED = "(" & ELED & ")" & "÷"
        End If
    End If
End If
If strcls = 1 Then

HLED = ""
ELED = ""
numone = "0"
strcls = 0
OldSign = ""
End If
HLED = HLED + CStr(Index)
ELED = ELED + CStr(Index)
Text1.Visible = False
EELED.Visible = True



End Sub
Private Sub cmpoint_Click()
If strcls = 1 Then
HLED = ""
ELED = ""
numone = "0"
strcls = 0
OldSign = ""
End If
If LED = "" Then
HLED = HLED + "0"
ELED = ELED + "0"
End If
'If Right(ELED, 1) <> "0" Then ELED = ELED + "0"
If InStr(HLED, ".") = 0 Then
HLED = HLED + "."
ELED = ELED + "."
End If
End Sub
Private Sub cmcls_Click()
HLED = ""
numone = "0"
ELED = ""
OldSign = ""
If strcls = 1 Then strcls = 0
Tip = "显示区里面已经清空了+_+"
Timer1 = True
End Sub
Private Sub cmleft_Click()
If Len(HLED) > 0 Then
HLED = Left(HLED, Len(HLED) - 1)
ELED = Left(ELED, Len(ELED) - 1)
End If
End Sub

'=====================================================================================================界面动画效果
Private Sub ccls()
If strup = 0 Then
    For nu = 0 To 9
        num(nu) = Form2.Image1(nu)
    Next
    For nus = 0 To 3
        sign(nus) = Form2.sign0(nus)
    Next
    cmcls = Form2.cmcls(0)
    cmleft = Form2.cmleft(0)
    enter = Form2.enter(0)
    cmpoint = Form2.cmpoint(0)
    cmsave = Form2.cmsave(0)
    cmzf = Form2.cmzf(0)
    cmcp = Form2.cmcp(0)
    cmecp = Form2.cmecp(0)
    new1 = Form2.new1(0)
    new2 = Form2.new2(0)
    new3 = Form2.new3(0)
    new4 = Form2.new4(0)
    closes = Form2.closes(0)
    mini = Form2.mini(0)
    cmabout = Form2.cmabout(0)
    Label4.ForeColor = &HFFFFFF
    Label3.ForeColor = &HFFFFFF
    
End If
End Sub
'==================================move事件
Private Sub EELED_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
End Sub
Private Sub LED_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
End Sub
Private Sub cmabout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmabout = Form2.cmabout(1)
End Sub
Private Sub new1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new1 = Form2.new1(1)
End Sub
Private Sub news1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new1 = Form2.new1(1)
End Sub

Private Sub mini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then mini = Form2.mini(1)
End Sub
Private Sub closes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then closes = Form2.closes(1)
End Sub
Private Sub new2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new2 = Form2.new2(1)
End Sub
Private Sub news2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new2 = Form2.new2(1)
End Sub

Private Sub new3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new3 = Form2.new3(1)
End Sub
Private Sub news3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new3 = Form2.new3(1)
End Sub

Private Sub new4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new4 = Form2.new4(1)
End Sub
Private Sub news4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then new4 = Form2.new4(1)
End Sub

Private Sub cmleft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmleft = Form2.cmleft(1)
End Sub
Private Sub cmcp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmcp = Form2.cmcp(1)
End Sub
Private Sub cmecp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmecp = Form2.cmecp(1)
End Sub
Private Sub cmzf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmzf = Form2.cmzf(1)
End Sub
Private Sub cmsave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmsave = Form2.cmsave(1)
End Sub
Private Sub enter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then enter = Form2.enter(1)
End Sub
Private Sub cmcls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmcls = Form2.cmcls(1)
End Sub
Private Sub cmpoint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmpoint = Form2.cmpoint(1)
End Sub
    

Private Sub sign_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ccls
    If strup = 0 Then
    sign(Index) = Form2.sign1(Index)
            'For nu = 0 To 3
             '   If nu <> Index Then
             '       sign(nu) = Form2.sign0(nu)
             '   End If
                
           ' Next
    End If
End Sub
Private Sub num_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
         ccls
    If strup = 0 Then
    num(Index) = Form2.Image2(Index)
           ' For nu = 0 To 9
           '     If nu <> Index Then
            '        num(nu) = Form2.Image1(nu)
            '    End If
           ' Next
    End If
End Sub
'==================================down事件
Private Sub cmcls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmcls = Form2.cmcls(2)
            strup = 1
End Sub
Private Sub cmzf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmzf = Form2.cmzf(2)
            strup = 1
End Sub
Private Sub enter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
             enter = Form2.enter(2)
            strup = 1
End Sub
Private Sub cmcp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmcp = Form2.cmcp(2)
            strup = 1
End Sub
Private Sub cmecp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmecp = Form2.cmecp(2)
            strup = 1
End Sub
Private Sub cmsave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
             cmsave = Form2.cmsave(2)
            strup = 1
End Sub
Private Sub cmpoint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmpoint = Form2.cmpoint(2)
            strup = 1
End Sub

Private Sub sign_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
            For nu = 0 To 3
            If nu = Index Then
                  sign(nu) = Form2.sign2(nu)
                End If
            Next
            strup = 1
End Sub
Private Sub num_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
            For nu = 0 To 9
            If nu = Index Then
                  num(nu) = Form2.Image3(nu)
                End If
            Next
            strup = 1
End Sub
Private Sub cmleft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmleft = Form2.cmleft(2)
            strup = 1
End Sub
'==================================up事件
Private Sub cmleft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub num_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub cmcls_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub sign_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub enter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub cmpoint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub cmsave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub cmzf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub cmcp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub cmecp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
'======================================================================================================
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF00&
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
Label4.ForeColor = &HFFFF&

End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
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


Private Sub Timer2_Timer()
If Dir(App.Path & "/old.tmp", vbDirectory) <> "" Then
Dim StringA As String
Open App.Path & "/old.tmp" For Input As #1
Input #1, StringA
Close #1
Kill App.Path & "/" & StringA
Kill App.Path & "\old.tmp"
End If
Timer2.Enabled = False
End Sub

Private Sub Tip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
End Sub
