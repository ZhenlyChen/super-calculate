VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00353535&
   BorderStyle     =   0  'None
   Caption         =   "单位转换"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   5790
   ScaleWidth      =   9360
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6720
      Top             =   480
   End
   Begin VB.ComboBox Com2 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox Com1 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ComboBox ComBo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   690
      ItemData        =   "Form5.frx":08CA
      Left            =   1440
      List            =   "Form5.frx":08E9
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "复制成功"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "复制详情"
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
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   4440
      Width           =   1935
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
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   120
      Picture         =   "Form5.frx":0923
      Stretch         =   -1  'True
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Tip 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "类型："
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
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   0
      Left            =   6720
      ToolTipText     =   "0"
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   1
      Left            =   6720
      ToolTipText     =   "1"
      Top             =   3960
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   2
      Left            =   7560
      ToolTipText     =   "2"
      Top             =   3960
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   3
      Left            =   8400
      ToolTipText     =   "3"
      Top             =   3960
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   4
      Left            =   6720
      ToolTipText     =   "4"
      Top             =   3120
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   5
      Left            =   7560
      ToolTipText     =   "5"
      Top             =   3120
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   6
      Left            =   8400
      ToolTipText     =   "6"
      Top             =   3120
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   7
      Left            =   6720
      ToolTipText     =   "7"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   8
      Left            =   7560
      ToolTipText     =   "8"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image num 
      Height          =   720
      Index           =   9
      Left            =   8400
      ToolTipText     =   "9"
      Top             =   2280
      Width           =   720
   End
   Begin VB.Image cmcls 
      Height          =   720
      Left            =   6720
      ToolTipText     =   "全部清空(Delete)"
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image cmleft 
      Height          =   735
      Left            =   8400
      ToolTipText     =   "退格"
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image cmpoint 
      Height          =   720
      Left            =   7560
      ToolTipText     =   "小数点"
      Top             =   4800
      Width           =   720
   End
   Begin VB.Label L2 
      BackColor       =   &H00404040&
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
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "单位转换"
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
      Height          =   855
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label L1 
      BackColor       =   &H00404040&
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Image closes 
      Height          =   480
      Left            =   8760
      Picture         =   "Form5.frx":11ED
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image min 
      Height          =   480
      Left            =   8160
      Picture         =   "Form5.frx":12E5
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strup As Integer
Dim cnu1 As Double
Dim cnu2 As Double
Private Sub closes_Click()
Unload Me
End Sub

Private Sub cmcls_Click()
strup = 0
L1 = ""
L2 = ""
End Sub

Private Sub cmleft_Click()
strup = 0
If Len(L1) > 0 Then
L1 = Left(L1, Len(L1) - 1)
End If
End Sub
Private Sub keyin(i As Integer)
ccls
If i = 229 Then MsgBox ("请关闭输入法或切换为英文状态再输入")
If 95 < Int(i) And Int(i) < 106 Then
    num_Click (Int(i) - 96)
            For nu = 0 To 9
                If nu = Int(Int(i) - 96) Then
                  num(nu) = Form2.Image3(nu)
                End If
            Next
    strup = 0
End If

If i = 8 Then
cmleft_Click
cmleft = Form2.cmleft(2)
End If

If i = 190 Or i = 110 Then
cmpoint_Click
cmpoint = Form2.cmpoint(2)
End If

If i = 46 Then
cmcls_Click
cmcls = Form2.cmcls(2)
End If

End Sub

Private Sub cmpoint_Click()
strup = 0
If L1 = "" Then L1 = L1 + "0"
If InStr(L1, ".") = 0 Then
L1 = L1 + "."
End If
End Sub


Private Sub Com1_Click()
If ComBo.ListIndex = 0 Then
    If Com1.ListIndex = 0 Then cnu1 = 1
    If Com1.ListIndex = 1 Then cnu1 = 1000
    If Com1.ListIndex = 2 Then cnu1 = 1000
    If Com1.ListIndex = 3 Then cnu1 = 0.001
    If Com1.ListIndex = 4 Then cnu1 = 0.035315
    If Com1.ListIndex = 5 Then cnu1 = 0.001308
    If Com1.ListIndex = 6 Then cnu1 = 202.8841
    If Com1.ListIndex = 7 Then cnu1 = 67.62804
    If Com1.ListIndex = 8 Then cnu1 = 33.81402
    If Com1.ListIndex = 9 Then cnu1 = 4.226753
    If Com1.ListIndex = 10 Then cnu1 = 2.113376
    If Com1.ListIndex = 11 Then cnu1 = 1.056688
    If Com1.ListIndex = 12 Then cnu1 = 0.264172
    If Com1.ListIndex = 13 Then cnu1 = 168.9364
    If Com1.ListIndex = 14 Then cnu1 = 56.31213
    If Com1.ListIndex = 15 Then cnu1 = 35.19508
    If Com1.ListIndex = 16 Then cnu1 = 1.759754
    If Com1.ListIndex = 17 Then cnu1 = 0.879877
    If Com1.ListIndex = 18 Then cnu1 = 0.219969
    If Com1.ListIndex = 19 Then cnu1 = 61.02374
    cnn
End If
If ComBo.ListIndex = 1 Then
    If Com1.ListIndex = 0 Then cnu1 = 1000000000
    If Com1.ListIndex = 1 Then cnu1 = 1000000
    If Com1.ListIndex = 2 Then cnu1 = 1000
    If Com1.ListIndex = 3 Then cnu1 = 100
    If Com1.ListIndex = 4 Then cnu1 = 1
    If Com1.ListIndex = 5 Then cnu1 = 0.001
    If Com1.ListIndex = 6 Then cnu1 = 39.37008
    If Com1.ListIndex = 7 Then cnu1 = 3.28084
    If Com1.ListIndex = 8 Then cnu1 = 1.093613
    If Com1.ListIndex = 9 Then cnu1 = 0.000621
    If Com1.ListIndex = 10 Then cnu1 = 0.00054
    cnn
End If
If ComBo.ListIndex = 2 Then
    If Com1.ListIndex = 0 Then cnu1 = 5
    If Com1.ListIndex = 1 Then cnu1 = 1000
    If Com1.ListIndex = 2 Then cnu1 = 100
    If Com1.ListIndex = 3 Then cnu1 = 10
    If Com1.ListIndex = 4 Then cnu1 = 1
    If Com1.ListIndex = 5 Then cnu1 = 0.1
    If Com1.ListIndex = 6 Then cnu1 = 0.01
    If Com1.ListIndex = 7 Then cnu1 = 0.001
    If Com1.ListIndex = 8 Then cnu1 = 0.000001
    If Com1.ListIndex = 9 Then cnu1 = 0.3527496
    If Com1.ListIndex = 10 Then cnu1 = 0.002205623
    If Com1.ListIndex = 11 Then cnu1 = 0.000157473
    If Com1.ListIndex = 12 Then cnu1 = 1.10231131092439E-06
    If Com1.ListIndex = 13 Then cnu1 = 9.84206527611061E-07
    If Com1.ListIndex = 14 Then cnu1 = 0.002
    If Com1.ListIndex = 15 Then cnu1 = 0.02
    If Com1.ListIndex = 16 Then cnu1 = 1.5747304441777E-04
    If Com1.ListIndex = 17 Then cnu1 = 0.56438339119329
    If Com1.ListIndex = 18 Then cnu1 = 15.432358352941
    If Com1.ListIndex = 19 Then cnu1 = 1.9684130552221E-05
    If Com1.ListIndex = 20 Then cnu1 = 2.2046226218488E-05
    cnn
End If
If ComBo.ListIndex = 3 Then
    If Com1.ListIndex = 0 Then cnu1 = 6.241509E+18
    If Com1.ListIndex = 1 Then cnu1 = 1
    If Com1.ListIndex = 2 Then cnu1 = 0.001
    If Com1.ListIndex = 3 Then cnu1 = 0.239005736137669
    If Com1.ListIndex = 4 Then cnu1 = 2.39005736137667E-04
    If Com1.ListIndex = 5 Then cnu1 = 0.737562149277266
    If Com1.ListIndex = 6 Then cnu1 = 9.4816987913438E-04
    cnn
End If
If ComBo.ListIndex = 4 Then
    If Com1.ListIndex = 0 Then cnu1 = 1000000
    If Com1.ListIndex = 1 Then cnu1 = 10000
    If Com1.ListIndex = 2 Then cnu1 = 1
    If Com1.ListIndex = 3 Then cnu1 = 0.0001
    If Com1.ListIndex = 4 Then cnu1 = 0.000001
    If Com1.ListIndex = 5 Then cnu1 = 1550.0031000062
    If Com1.ListIndex = 6 Then cnu1 = 10.7639104167097
    If Com1.ListIndex = 7 Then cnu1 = 1.19599004630108
    If Com1.ListIndex = 8 Then cnu1 = 2.47105381467165E-04
    If Com1.ListIndex = 9 Then cnu1 = 3.86102158542446E-07
    cnn
End If
If ComBo.ListIndex = 5 Then
    If Com1.ListIndex = 0 Then cnu1 = 100
    If Com1.ListIndex = 1 Then cnu1 = 1
    If Com1.ListIndex = 2 Then cnu1 = 3.6
    If Com1.ListIndex = 3 Then cnu1 = 3.28083989501312
    If Com1.ListIndex = 4 Then cnu1 = 2.23713646532438
    If Com1.ListIndex = 5 Then cnu1 = 1.94401244167963
    If Com1.ListIndex = 6 Then cnu1 = 2.9385836027035E-03
    cnn
End If
If ComBo.ListIndex = 6 Then
    If Com1.ListIndex = 0 Then cnu1 = 3600000000#
    If Com1.ListIndex = 1 Then cnu1 = 3600000
    If Com1.ListIndex = 2 Then cnu1 = 3600
    If Com1.ListIndex = 3 Then cnu1 = 60
    If Com1.ListIndex = 4 Then cnu1 = 1
    If Com1.ListIndex = 5 Then cnu1 = 0.0416666666667
    If Com1.ListIndex = 6 Then cnu1 = 5.95238095238095E-03
    If Com1.ListIndex = 7 Then cnu1 = 1.14077116130504E-04
    cnn
End If
If ComBo.ListIndex = 7 Then
    If Com1.ListIndex = 0 Then cnu1 = 1000
    If Com1.ListIndex = 1 Then cnu1 = 1
    If Com1.ListIndex = 2 Then cnu1 = 1.34102208959503
    If Com1.ListIndex = 3 Then cnu1 = 44253.728956636
    If Com1.ListIndex = 4 Then cnu1 = 56.8690192748062
    cnn
End If
If ComBo.ListIndex = 8 Then
    If Com1.ListIndex = 0 Then cnu1 = 1073741824
    If Com1.ListIndex = 1 Then cnu1 = 134217728
    If Com1.ListIndex = 2 Then cnu1 = 1048576
    If Com1.ListIndex = 3 Then cnu1 = 131072
    If Com1.ListIndex = 4 Then cnu1 = 1024
    If Com1.ListIndex = 5 Then cnu1 = 128
    If Com1.ListIndex = 6 Then cnu1 = 1
    If Com1.ListIndex = 7 Then cnu1 = 0.125
    If Com1.ListIndex = 8 Then cnu1 = 0.0009765625
    If Com1.ListIndex = 9 Then cnu1 = 0.0001220703125
    If Com1.ListIndex = 10 Then cnu1 = 9.5367431640625E-07
    If Com1.ListIndex = 11 Then cnu1 = 1.19209289550781E-07
   
    cnn
End If
End Sub
Private Sub Com2_Click()
If ComBo.ListIndex = 0 Then
    If Com2.ListIndex = 0 Then cnu2 = 1
    If Com2.ListIndex = 1 Then cnu2 = 1000
    If Com2.ListIndex = 2 Then cnu2 = 1000
    If Com2.ListIndex = 3 Then cnu2 = 0.001
    If Com2.ListIndex = 4 Then cnu2 = 0.035315
    If Com2.ListIndex = 5 Then cnu2 = 0.001308
    If Com2.ListIndex = 6 Then cnu2 = 202.8841
    If Com2.ListIndex = 7 Then cnu2 = 67.62804
    If Com2.ListIndex = 8 Then cnu2 = 33.81402
    If Com2.ListIndex = 9 Then cnu2 = 4.226753
    If Com2.ListIndex = 10 Then cnu2 = 2.113376
    If Com2.ListIndex = 11 Then cnu2 = 1.056688
    If Com2.ListIndex = 12 Then cnu2 = 0.264172
    If Com2.ListIndex = 13 Then cnu2 = 168.9364
    If Com2.ListIndex = 14 Then cnu2 = 56.31213
    If Com2.ListIndex = 15 Then cnu2 = 35.19508
    If Com2.ListIndex = 16 Then cnu2 = 1.759754
    If Com2.ListIndex = 17 Then cnu2 = 0.879877
    If Com2.ListIndex = 18 Then cnu2 = 0.219969
    If Com2.ListIndex = 19 Then cnu2 = 61.02374
    cnn
End If
If ComBo.ListIndex = 1 Then
    If Com2.ListIndex = 0 Then cnu2 = 1000000000
    If Com2.ListIndex = 1 Then cnu2 = 1000000
    If Com2.ListIndex = 2 Then cnu2 = 1000
    If Com2.ListIndex = 3 Then cnu2 = 100
    If Com2.ListIndex = 4 Then cnu2 = 1
    If Com2.ListIndex = 5 Then cnu2 = 0.001
    If Com2.ListIndex = 6 Then cnu2 = 39.37008
    If Com2.ListIndex = 7 Then cnu2 = 3.28084
    If Com2.ListIndex = 8 Then cnu2 = 1.093613
    If Com2.ListIndex = 9 Then cnu2 = 0.000621
    If Com2.ListIndex = 10 Then cnu2 = 0.00054
    cnn
End If
If ComBo.ListIndex = 2 Then
    If Com2.ListIndex = 0 Then cnu2 = 5
    If Com2.ListIndex = 1 Then cnu2 = 1000
    If Com2.ListIndex = 2 Then cnu2 = 100
    If Com2.ListIndex = 3 Then cnu2 = 10
    If Com2.ListIndex = 4 Then cnu2 = 1
    If Com2.ListIndex = 5 Then cnu2 = 0.1
    If Com2.ListIndex = 6 Then cnu2 = 0.01
    If Com2.ListIndex = 7 Then cnu2 = 0.001
    If Com2.ListIndex = 8 Then cnu2 = 0.000001
    If Com2.ListIndex = 9 Then cnu2 = 0.3527496
    If Com2.ListIndex = 10 Then cnu2 = 0.002205623
    If Com2.ListIndex = 11 Then cnu2 = 0.000157473
    If Com2.ListIndex = 12 Then cnu2 = 1.10231131092439E-06
    If Com2.ListIndex = 13 Then cnu2 = 9.84206527611061E-07
    If Com2.ListIndex = 14 Then cnu2 = 0.002
    If Com2.ListIndex = 15 Then cnu2 = 0.02
    If Com2.ListIndex = 16 Then cnu2 = 1.5747304441777E-04
    If Com2.ListIndex = 17 Then cnu2 = 0.56438339119329
    If Com2.ListIndex = 18 Then cnu2 = 15.432358352941
    If Com2.ListIndex = 19 Then cnu2 = 1.9684130552221E-05
    If Com2.ListIndex = 20 Then cnu2 = 2.2046226218488E-05
    cnn
End If
If ComBo.ListIndex = 3 Then
    If Com2.ListIndex = 0 Then cnu2 = 6.241509E+18
    If Com2.ListIndex = 1 Then cnu2 = 1
    If Com2.ListIndex = 2 Then cnu2 = 0.001
    If Com2.ListIndex = 3 Then cnu2 = 0.239005736137669
    If Com2.ListIndex = 4 Then cnu2 = 2.39005736137667E-04
    If Com2.ListIndex = 5 Then cnu2 = 0.737562149277266
    If Com2.ListIndex = 6 Then cnu2 = 9.4816987913438E-04
    cnn
End If
If ComBo.ListIndex = 4 Then
    If Com2.ListIndex = 0 Then cnu2 = 1000000
    If Com2.ListIndex = 1 Then cnu2 = 10000
    If Com2.ListIndex = 2 Then cnu2 = 1
    If Com2.ListIndex = 3 Then cnu2 = 0.0001
    If Com2.ListIndex = 4 Then cnu2 = 0.000001
    If Com2.ListIndex = 5 Then cnu2 = 1550.0031000062
    If Com2.ListIndex = 6 Then cnu2 = 10.7639104167097
    If Com2.ListIndex = 7 Then cnu2 = 1.19599004630108
    If Com2.ListIndex = 8 Then cnu2 = 2.47105381467165E-04
    If Com2.ListIndex = 9 Then cnu2 = 3.86102158542446E-07
    cnn
End If
If ComBo.ListIndex = 5 Then
    If Com2.ListIndex = 0 Then cnu2 = 100
    If Com2.ListIndex = 1 Then cnu2 = 1
    If Com2.ListIndex = 2 Then cnu2 = 3.6
    If Com2.ListIndex = 3 Then cnu2 = 3.28083989501312
    If Com2.ListIndex = 4 Then cnu2 = 2.23713646532438
    If Com2.ListIndex = 5 Then cnu2 = 1.94401244167963
    If Com2.ListIndex = 6 Then cnu2 = 2.9385836027035E-03
    cnn
End If
If ComBo.ListIndex = 6 Then
    If Com2.ListIndex = 0 Then cnu2 = 3600000000#
    If Com2.ListIndex = 1 Then cnu2 = 3600000
    If Com2.ListIndex = 2 Then cnu2 = 3600
    If Com2.ListIndex = 3 Then cnu2 = 60
    If Com2.ListIndex = 4 Then cnu2 = 1
    If Com2.ListIndex = 5 Then cnu2 = 0.0416666666667
    If Com2.ListIndex = 6 Then cnu2 = 5.95238095238095E-03
    If Com2.ListIndex = 7 Then cnu2 = 1.14077116130504E-04
    cnn
End If
If ComBo.ListIndex = 7 Then
    If Com2.ListIndex = 0 Then cnu2 = 1000
    If Com2.ListIndex = 1 Then cnu2 = 1
    If Com2.ListIndex = 2 Then cnu2 = 1.34102208959503
    If Com2.ListIndex = 3 Then cnu2 = 44253.728956636
    If Com2.ListIndex = 4 Then cnu2 = 56.8690192748062
    cnn
End If
If ComBo.ListIndex = 8 Then
    If Com2.ListIndex = 0 Then cnu2 = 1073741824
    If Com2.ListIndex = 1 Then cnu2 = 134217728
    If Com2.ListIndex = 2 Then cnu2 = 1048576
    If Com2.ListIndex = 3 Then cnu2 = 131072
    If Com2.ListIndex = 4 Then cnu2 = 1024
    If Com2.ListIndex = 5 Then cnu2 = 128
    If Com2.ListIndex = 6 Then cnu2 = 1
    If Com2.ListIndex = 7 Then cnu2 = 0.125
    If Com2.ListIndex = 8 Then cnu2 = 0.0009765625
    If Com2.ListIndex = 9 Then cnu2 = 0.0001220703125
    If Com2.ListIndex = 10 Then cnu2 = 9.5367431640625E-07
    If Com2.ListIndex = 11 Then cnu2 = 1.19209289550781E-07
   
    cnn
End If
End Sub
Private Sub Com1_KeyDown(KeyCode As Integer, Shift As Integer)
keyin (KeyCode)
End Sub
Private Sub Com2_KeyDown(KeyCode As Integer, Shift As Integer)
keyin (KeyCode)
End Sub
Private Sub ComBo_Click()

Com1.Clear
Com2.Clear
L1 = ""
L2 = ""
Tip = ""
    If ComBo.Text = "体积" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "升", 0
    Com1.AddItem "毫升", 1
    Com1.AddItem "立方厘米", 2
    Com1.AddItem "立方米", 3
    Com1.AddItem "立方英尺", 4
    Com1.AddItem "立方码", 5
    Com1.AddItem "茶匙（美）", 6
    Com1.AddItem "餐匙（美）", 7
    Com1.AddItem "液量盎司（美）", 8
    Com1.AddItem "杯（美）", 9
    Com1.AddItem "品脱（美）", 10
    Com1.AddItem "夸脱（美）", 11
    Com1.AddItem "加仑（美）", 12
    Com1.AddItem "茶匙（英）", 13
    Com1.AddItem "餐匙（英）", 14
    Com1.AddItem "液量盎司（英）", 15
    Com1.AddItem "品脱（英）", 16
    Com1.AddItem "夸脱（英）", 17
    Com1.AddItem "加仑（英）", 18
    Com1.AddItem "立方英寸", 19
    Com2.AddItem "升", 0
    Com2.AddItem "毫升", 1
    Com2.AddItem "立方厘米", 2
    Com2.AddItem "立方米", 3
    Com2.AddItem "立方英尺", 4
    Com2.AddItem "立方码", 5
    Com2.AddItem "茶匙（美）", 6
    Com2.AddItem "餐匙（美）", 7
    Com2.AddItem "液量盎司（美）", 8
    Com2.AddItem "杯（美）", 9
    Com2.AddItem "品脱（美）", 10
    Com2.AddItem "夸脱（美）", 11
    Com2.AddItem "加仑（美）", 12
    Com2.AddItem "茶匙（英）", 13
    Com2.AddItem "餐匙（英）", 14
    Com2.AddItem "液量盎司（英）", 15
    Com2.AddItem "品脱（英）", 16
    Com2.AddItem "夸脱（英）", 17
    Com2.AddItem "加仑（英）", 18
    Com2.AddItem "立方英寸", 19
    End If
    If ComBo.Text = "长度" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "纳米", 0
    Com1.AddItem "微米", 1
    Com1.AddItem "毫米", 2
    Com1.AddItem "厘米", 3
    Com1.AddItem "米", 4
    Com1.AddItem "公里", 5
    Com1.AddItem "英寸", 6
    Com1.AddItem "英尺", 7
    Com1.AddItem "码", 8
    Com1.AddItem "英里", 9
    Com1.AddItem "海里", 10
    Com2.AddItem "纳米", 0
    Com2.AddItem "微米", 1
    Com2.AddItem "毫米", 2
    Com2.AddItem "厘米", 3
    Com2.AddItem "米", 4
    Com2.AddItem "公里", 5
    Com2.AddItem "英寸", 6
    Com2.AddItem "英尺", 7
    Com2.AddItem "码", 8
    Com2.AddItem "英里", 9
    Com2.AddItem "海里", 10
    End If
    If ComBo.Text = "重量" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "克拉", 0
    Com1.AddItem "毫克", 1
    Com1.AddItem "厘克", 2
    Com1.AddItem "分克", 3
    Com1.AddItem "克", 4
    Com1.AddItem "十克", 5
    Com1.AddItem "百克", 6
    Com1.AddItem "千克", 7
    Com1.AddItem "公吨", 8
    Com1.AddItem "盎司", 9
    Com1.AddItem "磅", 10
    Com1.AddItem "石", 11
    Com1.AddItem "短吨", 12
    Com1.AddItem "长吨", 13
    Com1.AddItem "公斤", 14
    Com1.AddItem "两", 15
    Com1.AddItem "英石", 16
    Com1.AddItem "打兰", 17
    Com1.AddItem "格令", 18
    Com1.AddItem "英担", 19
    Com1.AddItem "美担", 20
    Com2.AddItem "克拉", 0
    Com2.AddItem "毫克", 1
    Com2.AddItem "厘克", 2
    Com2.AddItem "分克", 3
    Com2.AddItem "克", 4
    Com2.AddItem "十克", 5
    Com2.AddItem "百克", 6
    Com2.AddItem "千克", 7
    Com2.AddItem "公吨", 8
    Com2.AddItem "盎司", 9
    Com2.AddItem "磅", 10
    Com2.AddItem "石", 11
    Com2.AddItem "短吨", 12
    Com2.AddItem "长吨", 13
    Com2.AddItem "斤", 14
    Com2.AddItem "两", 15
    Com2.AddItem "英石", 16
    Com2.AddItem "打兰", 17
    Com2.AddItem "格令", 18
    Com2.AddItem "英担", 19
    Com2.AddItem "美担", 20
    End If
    If ComBo.Text = "能量" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "电子伏特", 0
    Com1.AddItem "焦耳", 1
    Com1.AddItem "千焦耳", 2
    Com1.AddItem "热量卡路里", 3
    Com1.AddItem "食物卡路里", 4
    Com1.AddItem "英尺-磅", 5
    Com1.AddItem "英制热量单位", 6
    Com2.AddItem "电子伏特", 0
    Com2.AddItem "焦耳", 1
    Com2.AddItem "千焦耳", 2
    Com2.AddItem "热量卡路里", 3
    Com2.AddItem "食物卡路里", 4
    Com2.AddItem "英尺-磅", 5
    Com2.AddItem "英制热量单位", 6
    End If
    If ComBo.Text = "面积" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "平方毫米", 0
    Com1.AddItem "平方厘米", 1
    Com1.AddItem "平方米", 2
    Com1.AddItem "公顷", 3
    Com1.AddItem "平方公里", 4
    Com1.AddItem "平方英寸", 5
    Com1.AddItem "平方英尺", 6
    Com1.AddItem "平方码", 7
    Com1.AddItem "英亩", 8
    Com1.AddItem "平方英里", 9
    Com2.AddItem "平方毫米", 0
    Com2.AddItem "平方厘米", 1
    Com2.AddItem "平方米", 2
    Com2.AddItem "公顷", 3
    Com2.AddItem "平方公里", 4
    Com2.AddItem "平方英寸", 5
    Com2.AddItem "平方英尺", 6
    Com2.AddItem "平方码", 7
    Com2.AddItem "英亩", 8
    Com2.AddItem "平方英里", 9
    End If
    If ComBo.Text = "速度" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "厘米/秒", 0
    Com1.AddItem "米/秒", 1
    Com1.AddItem "千米/小时", 2
    Com1.AddItem "英尺/秒", 3
    Com1.AddItem "英里/小时", 4
    Com1.AddItem "节", 5
    Com1.AddItem "马赫", 6
    Com2.AddItem "厘米/秒", 0
    Com2.AddItem "米/秒", 1
    Com2.AddItem "千米/小时", 2
    Com2.AddItem "英尺/秒", 3
    Com2.AddItem "英里/小时", 4
    Com2.AddItem "节", 5
    Com2.AddItem "马赫", 6
    End If
    If ComBo.Text = "时间" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "微秒", 0
    Com1.AddItem "毫秒", 1
    Com1.AddItem "秒", 2
    Com1.AddItem "分钟", 3
    Com1.AddItem "小时", 4
    Com1.AddItem "天", 5
    Com1.AddItem "周", 6
    Com1.AddItem "年", 7
    Com2.AddItem "微秒", 0
    Com2.AddItem "毫秒", 1
    Com2.AddItem "秒", 2
    Com2.AddItem "分钟", 3
    Com2.AddItem "小时", 4
    Com2.AddItem "天", 5
    Com2.AddItem "周", 6
    Com2.AddItem "年", 7
    End If
    If ComBo.Text = "功率" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "瓦特", 0
    Com1.AddItem "千瓦", 1
    Com1.AddItem "马力（美）", 2
    Com1.AddItem "英尺-磅/分钟", 3
    Com1.AddItem "英制热量单位/分钟", 4
    Com2.AddItem "瓦特", 0
    Com2.AddItem "千瓦", 1
    Com2.AddItem "马力（美）", 2
    Com2.AddItem "英尺-磅/分钟", 3
    Com2.AddItem "英制热量单位/分钟", 4
    End If
    If ComBo.Text = "数据" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "位", 0
    Com1.AddItem "字节", 1
    Com1.AddItem "千位", 2
    Com1.AddItem "千字节", 3
    Com1.AddItem "兆位", 4
    Com1.AddItem "兆字节", 5
    Com1.AddItem "千兆位", 6
    Com1.AddItem "千兆字节", 7
    Com1.AddItem "千吉位", 8
    Com1.AddItem "千吉字节", 9
    Com1.AddItem "万兆位", 10
    Com1.AddItem "万兆字节", 11
    Com2.AddItem "位", 0
    Com2.AddItem "字节", 1
    Com2.AddItem "千位", 2
    Com2.AddItem "千字节", 3
    Com2.AddItem "兆位", 4
    Com2.AddItem "兆字节", 5
    Com2.AddItem "千兆位", 6
    Com2.AddItem "千兆字节", 7
    Com2.AddItem "千吉位", 8
    Com2.AddItem "千吉字节", 9
    Com2.AddItem "万兆位", 10
    Com2.AddItem "万兆字节", 11
    End If
    cnn

    End Sub

Private Sub ComBo_KeyDown(KeyCode As Integer, Shift As Integer)
keyin (KeyCode)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
keyin (KeyCode)
End Sub

Private Sub Form_Load()
    strup = 0
End Sub


Private Sub cnn()
Dim num1, num2, num3, num4 As String
If L1 <> "" Then
    If Right(L1, 1) <> "." Then
        If Com1.Text <> "" And Com2.Text <> "" Then
            L2 = CStr(CSng(L1) / cnu1 * cnu2)
            If Left(L2, 1) = "." Then L2 = "0" & L2
            
            
            If ComBo.ListIndex = 0 Then
                num1 = Format(CDbl(L1) / cnu1 * 0.035315, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 0.219969, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 1.759754, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  立方英尺" & vbCrLf & num2 & "  加仑(英)   " & vbCrLf & num3 & "  品脱(英)   "
                    If Format(num4 * 0.00000026667, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.00000026667, "0.00")) & "  个游泳池"
                    Else
                        If Format(num4 * 0.002642, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.002642, "0.00")) & "  个浴缸"
                            Else
                                If Format(num4 * 4.226753, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 4.226753, "0.00")) & "  个咖啡杯"
                        End If
                    End If
            End If
            
            
            If ComBo.ListIndex = 1 Then
                num1 = Format(CDbl(L1) / cnu1 * 39.37008, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 3.28084, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 100, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  英寸" & vbCrLf & num2 & "  英尺" & vbCrLf & num3 & "  厘米"
                    If Format(num4 * 0.01316, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.01316, "0.00")) & "  架大型喷气式客机"
                    Else
                        If Format(num4 * 5.351351, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 5.351351, "0.00")) & "  只手"
                            Else
                                If Format(num4 * 28.545454, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 28.545454, "0.00")) & "  个曲别针"
                        End If
                    End If
            End If
            
            
            
            If ComBo.ListIndex = 2 Then
                num1 = Format(CDbl(L1) / cnu1 * 0.001, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 0.002205623, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 0.3527496, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  千克" & vbCrLf & num2 & "  磅" & vbCrLf & num3 & "  盎司"
                 If Format(num4 * 0.0000000111, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0000000111, "0.00")) & "  只鲸"
                    Else
                    If Format(num4 * 0.00000025, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.00000025, "0.00")) & "  头大象"
                    Else
                        If Format(num4 * 0.002312, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.002312, "0.00")) & "  个足球"
                            Else
                                If Format(num4 * 500, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 500, "0.00")) & "  片雪花"
                        End If
                    End If
                    End If
            End If
            
            If ComBo.ListIndex = 3 Then
                num1 = Format(CDbl(L1) / cnu1, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 0.239005736137669, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 9.4816987913438E-04, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  焦耳" & vbCrLf & num2 & "  千卡" & vbCrLf & num3 & "  BTU"

                    If Format(num4 * 0.0000009554, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0000009554, "0.00")) & "  只蛋糕"
                    Else
                        If Format(num4 * 0.0000023, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0000023, "0.00")) & "  只香蕉"
                            Else
                                If Format(num4 * 0.000111, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 0.000111, "0.00")) & "  节电池"
                        End If
                    End If
             End If
                    

            If ComBo.ListIndex = 4 Then
                num1 = Format(CDbl(L1) / cnu1 * 0.0001, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 2.47105381467165E-04, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 1.19599004630108, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  公顷" & vbCrLf & num2 & "  英亩" & vbCrLf & num3 & "  平方码"

                    If Format(num4 * 0.00001, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.00001, "0.00")) & "  个城堡"
                    Else
                        If Format(num4 * 16.58, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 16.58, "0.00")) & "  张纸"
                            Else
                                If Format(num4 * 80, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 80, "0.00")) & "  只手"
                        End If
                    End If
             End If
             
             
            If ComBo.ListIndex = 5 Then
                num1 = Format(CDbl(L1) / cnu1 * 1, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 2.9385836027035E-03, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 2.23713646532438, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  米/秒" & vbCrLf & num2 & "  马赫" & vbCrLf & num3 & "  英里/小时"

                    If Format(num4 * 0.004068, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.004068, "0.00")) & "  架喷气式飞机"
                    Else
                        If Format(num4 * 0.0498, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0498, "0.00")) & "  匹马"
                            Else
                                If Format(num4 * 11.19, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 11.19, "0.00")) & "  只乌龟"
                        End If
                    End If
             End If
             
            If ComBo.ListIndex = 6 Then
                num1 = Format(CDbl(L1) / cnu1 * 3600000, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 0.0416666666667, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 1.14077116130504E-04, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  秒" & vbCrLf & num2 & "  天" & vbCrLf & num3 & "  年"
             End If
             
            If ComBo.ListIndex = 7 Then
                num1 = Format(CDbl(L1) / cnu1 * 3600000000#, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 1.34102208959503, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 56.8690192748062, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  瓦特" & vbCrLf & num2 & "  马力" & vbCrLf & num3 & "  BTU/分钟"

                    If Format(num4 * 0.0003352666, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0003352666, "0.00")) & "  个火车引擎"
                    Else
                        If Format(num4 * 1.3413333, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 1.3413333, "0.00")) & "  匹马"
                            Else
                                If Format(num4 * 16.6455696203, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 16.6455696203, "0.00")) & "  个灯泡"
                        End If
                    End If
             End If
             
            If ComBo.ListIndex = 8 Then
                num1 = Format(CDbl(L1) / cnu1 * 1024, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 134217728, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 1, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "还不到1"
                If num2 = "0.00" Then num2 = "还不到1"
                If num3 = "0.00" Then num3 = "还不到1"
                Tip = "约等于" & vbCrLf & num1 & "  兆位" & vbCrLf & num2 & "  字节" & vbCrLf & num3 & "  千兆位"

                    If Format(num4 * 0.026666667, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.026666667, "0.00")) & "  张DVD"
                    Else
                        If Format(num4 * 0.183, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.183, "0.00")) & "  张CD"
                            Else
                                If Format(num4 * 88.927943761, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 88.927943761, "0.00")) & "  张软盘"
                        End If
                    End If
             End If
             
             
             
        End If
    End If
Else
L2 = ""
Tip = ""
End If
End Sub

Private Sub L1_Change()
cnn
End Sub

Private Sub Label3_Click()
Clipboard.Clear
Clipboard.SetText L2
Label5.Visible = True
Timer1.Enabled = True
End Sub
Private Sub Label4_Click()
Clipboard.Clear
Clipboard.SetText L1 & Com1.Text & "=" & L2 & Com2.Text & vbCrLf & Tip.Caption
Label5.Visible = True
Timer1.Enabled = True
End Sub
Private Sub min_Click()
Me.WindowState = 1
End Sub
Private Sub num_Click(Index As Integer)

strup = 0
If Len(L1) < 15 Then L1 = L1 & CStr(Index)
End Sub
'=====================================================================================================界面动画效果
Private Sub ccls()
If strup = 0 Then
    For nu = 0 To 9
        num(nu) = Form2.Image1(nu)
    Next
    cmcls = Form2.cmcls(0)
    cmleft = Form2.cmleft(0)
    cmpoint = Form2.cmpoint(0)
    closes = Form2.closes(0)
    min = Form2.min(0)
    Label3.ForeColor = &HFFFFFF
    Label4.ForeColor = &HFFFFFF
End If
End Sub
'==================================move事件
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
End Sub
Private Sub closes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closes = Form2.closes(1)
min = Form2.min(0)
End Sub
Private Sub min_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closes = Form2.closes(0)
min = Form2.min(1)
End Sub
Private Sub cmleft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmleft = Form2.cmleft(1)
End Sub
Private Sub cmcls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmcls = Form2.cmcls(1)
End Sub
Private Sub cmpoint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then cmpoint = Form2.cmpoint(1)
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

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then Label3.ForeColor = &HFFFF&
End Sub



Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ccls
If strup = 0 Then Label4.ForeColor = &HFFFF&
End Sub


'==================================down事件
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF00&
strup = 1
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF00&
strup = 1
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub cmcls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmcls = Form2.cmcls(2)
            strup = 1
End Sub

Private Sub cmpoint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            cmpoint = Form2.cmpoint(2)
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
Private Sub cmpoint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strup = 0
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFFFF
strup = 0
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
strup = 0
End Sub
'======================================================================================================



Private Sub Timer1_Timer()
Label5.Visible = False
Timer1.Enabled = False
End Sub
