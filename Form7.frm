VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "函数模式 Beta1.3"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6795
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command9 
      BackColor       =   &H8000000A&
      Caption         =   "几何画板"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text17 
      Height          =   2055
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   56
      Text            =   "Form7.frx":08CA
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "分析结果"
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      TabIndex        =   21
      Top             =   2040
      Width           =   6735
      Begin VB.CommandButton cClear 
         BackColor       =   &H8000000A&
         Caption         =   "清空"
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4080
         Width           =   2055
      End
      Begin VB.PictureBox Pic 
         BackColor       =   &H00404040&
         Height          =   4215
         Left            =   120
         ScaleHeight     =   4155
         ScaleWidth      =   4155
         TabIndex        =   22
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   4215
         Left            =   4440
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "二次函数"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2400
         TabIndex        =   55
         Top             =   435
         Width           =   375
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   54
         Top             =   435
         Width           =   495
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   600
         TabIndex        =   53
         Top             =   435
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H8000000A&
         Caption         =   "分析"
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "y=         x^2+         x+"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "一次函数"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   0
      Left            =   2520
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "分析"
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1800
         TabIndex        =   19
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "y=                 x+"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "函数类型"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "直线方程"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   7
         Left            =   1320
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "对数函数"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "指数函数"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "三角函数"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "反比函数"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "三次函数"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "二次函数"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton obChoice 
         BackColor       =   &H00404040&
         Caption         =   "一次函数"
         ForeColor       =   &H8000000B&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "反比函数1"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   3
      Left            =   2520
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   600
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000A&
         Caption         =   "分析"
         Height          =   495
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "x"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "y=————————"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "直线方程"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   7
      Left            =   2520
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "研发中。。。。。。      敬请期待                       【关于--检查更新获取最新版本了解最新功能】"
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "对数函数"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   6
      Left            =   2520
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text13 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   49
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H8000000A&
         Caption         =   "分析"
         Height          =   495
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "y=log(                     )x"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   48
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "指数函数"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   5
      Left            =   2520
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   600
         TabIndex        =   46
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H8000000A&
         Caption         =   "分析"
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "y=                          ^x"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "三角函数"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   4
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   43
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1440
         TabIndex        =   41
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   40
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H8000000A&
         Caption         =   "分析"
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton Op 
         BackColor       =   &H00404040&
         Caption         =   "tan"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Op 
         BackColor       =   &H00404040&
         Caption         =   "cos"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   840
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Op 
         BackColor       =   &H00404040&
         Caption         =   "sin"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "y=         *cos(                    x+              )+    "
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H00404040&
      Caption         =   "三次函数"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Index           =   2
      Left            =   2520
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   33
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   32
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   480
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000A&
         Caption         =   "分析"
         Height          =   420
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "y=           x^3+          x^2+             x+ "
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim k As Double
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim X As Single
    Dim Y As Single
    Dim ibTitle As String, vColor As Long
     Const StepValue As Single = 0.001
  Dim mChoice As Byte
  Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwflags As Long, ByVal dwReserved As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


 

Public Function GetConnectionString() As String
  Dim dwflags As Long
  Dim MyMsg As String
  '显示的信息字符串
     If InternetGetConnectedState(dwflags, 0&) Then
        If dwflags And INTERNET_CONNECTION_CONFIGURED Then
           MyMsg = MyMsg + "系统配置了网络连接" + vbCrLf
        End If
        If dwflags And INTERNET_CONNECTION_LAN Then
           MyMsg = MyMsg + "系统通过路由器与Internet连接"
        End If
        If dwflags And INTERNET_CONNECTION_PROXY Then
           MyMsg = MyMsg + "并使用了Proxy代理服务器"
        Else
           MyMsg = MyMsg + "。"
        End If
        If dwflags And INTERNET_CONNECTION_MODEM Then
           MyMsg = MyMsg + "系统使用Modem与Internet连接"
        End If
        If dwflags And INTERNET_CONNECTION_OFFLINE Then
           MyMsg = MyMsg + "系统当前处于离线状态"
        End If
        If dwflags And INTERNET_CONNECTION_MODEM_BUSY Then
           MyMsg = MyMsg + "系统通ADSL连接到网络"
        End If
        If dwflags And INTERNET_RAS_INSTALLED Then
           MyMsg = MyMsg + "系统安装了远程访问服务"
        End If
     Else
        MyMsg = "系统当前未与Internet连接"
     End If
     GetConnectionString = MyMsg
End Function


    Private Sub dims()
   

            Pic.AutoRedraw = False
    vColor = RGB(255, 23, 34)
    End Sub
  

Sub Draw(Form As PictureBox, X As Integer, Y As Integer)
    Const Offset As Single = 0.5
    Dim i As Long
    Form.AutoRedraw = True
    Form.Scale (-X, Y)-(X, -Y)  '定义坐标系吖
    Form.Line (-X, 0)-(X, 0)    'X轴
    Form.Line (0, -Y)-(0, Y)    'Y轴
    For i = -X To X - 1
        Form.Line (i, 0)-(i, 0.2)   'X轴点
    Next i
    For i = -Y To Y - 1
        Form.Line (0, i)-(0.2, i)   'Y轴点
    Next i
    Form.Line (X, 0)-(X - Offset, Offset)      'X箭头
    Form.Line (X, 0)-(X - Offset, -Offset)
    Form.Line (0, Y)-(-Offset, Y - Offset)     'Y箭头
    Form.Line (0, Y)-(Offset, Y - Offset)
End Sub
Private Sub cClear_Click()
    Pic.Cls
    Draw Pic, 10, 10
End Sub


Private Sub Command1_Click()
dims
If Text1 <> "" And Text2 <> "" Then
            k = Text1
            b = Text2
            For X = -10 To 10 Step StepValue
                Y = k * X + b
                Pic.PSet (X, Y), vColor
                DoEvents
            Next
            Label2 = "一次函数" + vbCrLf + "与y轴交点：(0," + CStr(b) + ")" + vbCrLf + "与x轴交点：(" + Format(CStr(b / k * -1), "0.00") + ",0)"
End If

End Sub

Private Sub Command2_Click()
dims
If Text3 <> "" Then
            k = Text3
            For X = -10 To 10 Step StepValue
                Y = k / X
                Pic.PSet (X, Y), vColor
                DoEvents
            Next
            Label2 = "反比例函数" + vbCrLf + "任意坐标与xy轴平行线构成矩形面积：" & k
End If
End Sub

Private Sub Command3_Click()
dims

If Text4 <> "" And Text5 <> "" And Text6 <> "" And Text7 <> "" Then
            a = Text4
            b = Text5
            c = Text6
            c = Text7
            For X = -10 To 10 Step StepValue
                Y = a * (X ^ 3) + b * X ^ 2 + c * X + d
                Pic.PSet (X, Y), vColor
                DoEvents
            Next
            Label2 = "三次函数" + vbCrLf

End If
End Sub

Private Sub Command4_Click()
dims
If Text9 <> "" And Text8 <> "" And Text10 <> "" And Text11 <> "" Then
k = Text9
            a = Text8
            b = Text10
            c = Text11
            If Op(1).Value = True Then
                    For X = -10 To 10 Step StepValue
                        Y = a * Sin(k * X + c) + b
                        Pic.PSet (X, Y), vColor
                        DoEvents
                    Next
              End If
              If Op(2).Value = True Then
                    For X = -10 To 10 Step StepValue
                        Y = a * Cos(k * X + c) + b
                        Pic.PSet (X, Y), vColor
                        DoEvents
                    Next
              End If
              If Op(0).Value = True Then
                    For X = -10 To 10 Step StepValue
                        Y = a * Tan(k * X + c) + b
                        Pic.PSet (X, Y), vColor
                        DoEvents
                    Next
            End If
            Label2 = "三角函数" + vbCrLf

            End If
End Sub

Private Sub Command5_Click()
dims
If Text12 <> "" Then
            a = Text12
            For X = -10 To 10 Step StepValue
            
                Y = a ^ X
                
                If Y > 10 Then GoTo finished
                Pic.PSet (X, Y), vColor
                DoEvents
            Next
            Label2 = "指数函数" + vbCrLf
            End If
            Exit Sub
        
finished:
End Sub

Private Sub Command6_Click()
dims
If Text13 <> "" And Text13 <> "1" Then
            a = Text13
            For X = StepValue To 10 Step StepValue
                Y = Log(X) / Log(a)
                Pic.PSet (X, Y), vColor
                DoEvents
            Next
            Label2 = "对数函数" + vbCrLf
            End If
End Sub

Private Sub Command7_Click()
dims
If Text14 <> "" And Text15 <> "" And Text16 <> "" Then
            a = Text14
            b = Text15
            c = Text16
            For X = -10 To 10 Step StepValue
                Y = a * (X ^ 2) + b * X + c
                Pic.PSet (X, Y), vColor
                DoEvents
            Next
            Label2 = "二次函数" + vbCrLf + "对称轴：" + vbCrLf + "直线x=" + CStr(-1 * b / (2 * a)) + vbCrLf
            


If b * b - a * 4 * c > 0 Then Label2 = Label2 + "与x轴交点：" + vbCrLf + "（" + Format(CStr(-1 * b + Sqr(b * b - a * 4 * c) / (2 * a)), "0.00") + "," + Format(CStr(a * (-1 * b - Sqr(b * b - a * 4 * c) / (2 * a) ^ 2) + b * -1 * b * Sqr(b * b - a * 4 * c) / (2 * a) + c), "0.00") + ")" + vbCrLf + "与x轴交点：" + vbCrLf + "（" + Format(CStr(-1 * b - Sqr(b * b - a * 4 * c) / (2 * a)), "0.00") + "," + Format(CStr(a * (-1 * b + Sqr(b * b - a * 4 * c) / (2 * a) ^ 2) + b * -1 * b * Sqr(b * b - a * 4 * c) / (2 * a) + c), "0.00") + ")" + vbCrLf
If b * b - a * 4 * c = 0 Then Label2 = Label2 + "与x轴交点：" + vbCrLf + "（" + Format(CStr(-1 * b / (2 * a)), "0.00") + ",0)" + vbCrLf
If b * b - a * 4 * c < 0 Then Label2 = Label2 + "与x轴没有交点" + vbCrLf
End If
End Sub

Private Sub Command8_Click()

 If InternetCheckConnection("http://www.baidu.com/", 1, 0) = 0 Then
     MsgBox "网络不正常,无法连接服务器" + vbCrLf + "详细信息：" + GetConnectionString
    Else
Form8.Show
End If
End Sub

Private Sub Command9_Click()
If Dir(App.Path & "\几何画板.exe", vbDirectory) <> "" Then
    Open App.Path & "\start.bat" For Output As #1
    Print #1, "start " & App.Path & "\几何画板.exe"
    Close
    Shell ("cmd.exe /c " & App.Path & "\start.bat")
   ' Kill App.Path & "\start.bat"
    
Else
    msg = MsgBox("系统检测到你没有安装几何画板，是否立刻下载？" + vbCrLf + StringB, vbYesNo, "提示")
    If msg = vbYes Then
    Open App.Path & "\start.vbs" For Output As #1
    Print #1, Text17
    Close
    Shell "wscript.exe " & App.Path & "\start.vbs"""
    End If
End If

End Sub

Private Sub Form_Load()
    cClear_Click

End Sub

Private Sub obChoice_Click(Index As Integer)
fshow (Index)
  
End Sub
 
Private Sub fshow(a As Integer)
For i = 0 To 7
Fr(i).Visible = False
Next
Fr(a).Visible = True
End Sub
Private Sub fen()


End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
    Case 45
    Case 46
    Case 9
    Case 48 To 57 '数字
    If Len(Text1) - InStr(Text1, ".") > 4 And InStr(Text1, ".") > 0 Then KeyAscii = 0

    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text2) - InStr(Text2, ".") > 4 And InStr(Text2, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text3) - InStr(Text3, ".") > 4 And InStr(Text3, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text4) - InStr(Text4, ".") > 4 And InStr(Text4, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text5) - InStr(Text5, ".") > 4 And InStr(Text5, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text6) - InStr(Text6, ".") > 4 And InStr(Text6, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text7) - InStr(Text7, ".") > 4 And InStr(Text7, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text8) - InStr(Text8, ".") > 4 And InStr(Text8, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text9) - InStr(Text9, ".") > 4 And InStr(Text9, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text10) - InStr(Text10, ".") > 4 And InStr(Text10, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text11) - InStr(Text11, ".") > 4 And InStr(Text11, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text12) - InStr(Text12, ".") > 4 And InStr(Text12, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text13) - InStr(Text13, ".") > 4 And InStr(Text13, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text14) - InStr(Text14, ".") > 4 And InStr(Text14, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub Text15_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text15) - InStr(Text15, ".") > 4 And InStr(Text15, ".") > 0 Then KeyAscii = 0
    Case Else
      KeyAscii = 0
  End Select
 
    
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8  '退格键
     
    Case 9
     
    Case 46
    Case 45
    Case 48 To 57 '数字
     If Len(Text16) - InStr(Text16, ".") > 4 And InStr(Text16, ".") > 0 Then KeyAscii = 0
    
    Case Else
      KeyAscii = 0
  End Select
End Sub
