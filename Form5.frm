VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00353535&
   BorderStyle     =   0  'None
   Caption         =   "��λת��"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   2  '��Ļ����
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "���Ƴɹ�"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "���ƽ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "���ͣ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      ToolTipText     =   "ȫ�����(Delete)"
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image cmleft 
      Height          =   735
      Left            =   8400
      ToolTipText     =   "�˸�"
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image cmpoint 
      Height          =   720
      Left            =   7560
      ToolTipText     =   "С����"
      Top             =   4800
      Width           =   720
   End
   Begin VB.Label L2 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��λת��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      ToolTipText     =   "�ر�"
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
If i = 229 Then MsgBox ("��ر����뷨���л�ΪӢ��״̬������")
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
    If ComBo.Text = "���" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "��", 0
    Com1.AddItem "����", 1
    Com1.AddItem "��������", 2
    Com1.AddItem "������", 3
    Com1.AddItem "����Ӣ��", 4
    Com1.AddItem "������", 5
    Com1.AddItem "��ף�����", 6
    Com1.AddItem "�ͳף�����", 7
    Com1.AddItem "Һ����˾������", 8
    Com1.AddItem "��������", 9
    Com1.AddItem "Ʒ�ѣ�����", 10
    Com1.AddItem "���ѣ�����", 11
    Com1.AddItem "���أ�����", 12
    Com1.AddItem "��ף�Ӣ��", 13
    Com1.AddItem "�ͳף�Ӣ��", 14
    Com1.AddItem "Һ����˾��Ӣ��", 15
    Com1.AddItem "Ʒ�ѣ�Ӣ��", 16
    Com1.AddItem "���ѣ�Ӣ��", 17
    Com1.AddItem "���أ�Ӣ��", 18
    Com1.AddItem "����Ӣ��", 19
    Com2.AddItem "��", 0
    Com2.AddItem "����", 1
    Com2.AddItem "��������", 2
    Com2.AddItem "������", 3
    Com2.AddItem "����Ӣ��", 4
    Com2.AddItem "������", 5
    Com2.AddItem "��ף�����", 6
    Com2.AddItem "�ͳף�����", 7
    Com2.AddItem "Һ����˾������", 8
    Com2.AddItem "��������", 9
    Com2.AddItem "Ʒ�ѣ�����", 10
    Com2.AddItem "���ѣ�����", 11
    Com2.AddItem "���أ�����", 12
    Com2.AddItem "��ף�Ӣ��", 13
    Com2.AddItem "�ͳף�Ӣ��", 14
    Com2.AddItem "Һ����˾��Ӣ��", 15
    Com2.AddItem "Ʒ�ѣ�Ӣ��", 16
    Com2.AddItem "���ѣ�Ӣ��", 17
    Com2.AddItem "���أ�Ӣ��", 18
    Com2.AddItem "����Ӣ��", 19
    End If
    If ComBo.Text = "����" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "����", 0
    Com1.AddItem "΢��", 1
    Com1.AddItem "����", 2
    Com1.AddItem "����", 3
    Com1.AddItem "��", 4
    Com1.AddItem "����", 5
    Com1.AddItem "Ӣ��", 6
    Com1.AddItem "Ӣ��", 7
    Com1.AddItem "��", 8
    Com1.AddItem "Ӣ��", 9
    Com1.AddItem "����", 10
    Com2.AddItem "����", 0
    Com2.AddItem "΢��", 1
    Com2.AddItem "����", 2
    Com2.AddItem "����", 3
    Com2.AddItem "��", 4
    Com2.AddItem "����", 5
    Com2.AddItem "Ӣ��", 6
    Com2.AddItem "Ӣ��", 7
    Com2.AddItem "��", 8
    Com2.AddItem "Ӣ��", 9
    Com2.AddItem "����", 10
    End If
    If ComBo.Text = "����" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "����", 0
    Com1.AddItem "����", 1
    Com1.AddItem "���", 2
    Com1.AddItem "�ֿ�", 3
    Com1.AddItem "��", 4
    Com1.AddItem "ʮ��", 5
    Com1.AddItem "�ٿ�", 6
    Com1.AddItem "ǧ��", 7
    Com1.AddItem "����", 8
    Com1.AddItem "��˾", 9
    Com1.AddItem "��", 10
    Com1.AddItem "ʯ", 11
    Com1.AddItem "�̶�", 12
    Com1.AddItem "����", 13
    Com1.AddItem "����", 14
    Com1.AddItem "��", 15
    Com1.AddItem "Ӣʯ", 16
    Com1.AddItem "����", 17
    Com1.AddItem "����", 18
    Com1.AddItem "Ӣ��", 19
    Com1.AddItem "����", 20
    Com2.AddItem "����", 0
    Com2.AddItem "����", 1
    Com2.AddItem "���", 2
    Com2.AddItem "�ֿ�", 3
    Com2.AddItem "��", 4
    Com2.AddItem "ʮ��", 5
    Com2.AddItem "�ٿ�", 6
    Com2.AddItem "ǧ��", 7
    Com2.AddItem "����", 8
    Com2.AddItem "��˾", 9
    Com2.AddItem "��", 10
    Com2.AddItem "ʯ", 11
    Com2.AddItem "�̶�", 12
    Com2.AddItem "����", 13
    Com2.AddItem "��", 14
    Com2.AddItem "��", 15
    Com2.AddItem "Ӣʯ", 16
    Com2.AddItem "����", 17
    Com2.AddItem "����", 18
    Com2.AddItem "Ӣ��", 19
    Com2.AddItem "����", 20
    End If
    If ComBo.Text = "����" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "���ӷ���", 0
    Com1.AddItem "����", 1
    Com1.AddItem "ǧ����", 2
    Com1.AddItem "������·��", 3
    Com1.AddItem "ʳ�￨·��", 4
    Com1.AddItem "Ӣ��-��", 5
    Com1.AddItem "Ӣ��������λ", 6
    Com2.AddItem "���ӷ���", 0
    Com2.AddItem "����", 1
    Com2.AddItem "ǧ����", 2
    Com2.AddItem "������·��", 3
    Com2.AddItem "ʳ�￨·��", 4
    Com2.AddItem "Ӣ��-��", 5
    Com2.AddItem "Ӣ��������λ", 6
    End If
    If ComBo.Text = "���" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "ƽ������", 0
    Com1.AddItem "ƽ������", 1
    Com1.AddItem "ƽ����", 2
    Com1.AddItem "����", 3
    Com1.AddItem "ƽ������", 4
    Com1.AddItem "ƽ��Ӣ��", 5
    Com1.AddItem "ƽ��Ӣ��", 6
    Com1.AddItem "ƽ����", 7
    Com1.AddItem "ӢĶ", 8
    Com1.AddItem "ƽ��Ӣ��", 9
    Com2.AddItem "ƽ������", 0
    Com2.AddItem "ƽ������", 1
    Com2.AddItem "ƽ����", 2
    Com2.AddItem "����", 3
    Com2.AddItem "ƽ������", 4
    Com2.AddItem "ƽ��Ӣ��", 5
    Com2.AddItem "ƽ��Ӣ��", 6
    Com2.AddItem "ƽ����", 7
    Com2.AddItem "ӢĶ", 8
    Com2.AddItem "ƽ��Ӣ��", 9
    End If
    If ComBo.Text = "�ٶ�" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "����/��", 0
    Com1.AddItem "��/��", 1
    Com1.AddItem "ǧ��/Сʱ", 2
    Com1.AddItem "Ӣ��/��", 3
    Com1.AddItem "Ӣ��/Сʱ", 4
    Com1.AddItem "��", 5
    Com1.AddItem "���", 6
    Com2.AddItem "����/��", 0
    Com2.AddItem "��/��", 1
    Com2.AddItem "ǧ��/Сʱ", 2
    Com2.AddItem "Ӣ��/��", 3
    Com2.AddItem "Ӣ��/Сʱ", 4
    Com2.AddItem "��", 5
    Com2.AddItem "���", 6
    End If
    If ComBo.Text = "ʱ��" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "΢��", 0
    Com1.AddItem "����", 1
    Com1.AddItem "��", 2
    Com1.AddItem "����", 3
    Com1.AddItem "Сʱ", 4
    Com1.AddItem "��", 5
    Com1.AddItem "��", 6
    Com1.AddItem "��", 7
    Com2.AddItem "΢��", 0
    Com2.AddItem "����", 1
    Com2.AddItem "��", 2
    Com2.AddItem "����", 3
    Com2.AddItem "Сʱ", 4
    Com2.AddItem "��", 5
    Com2.AddItem "��", 6
    Com2.AddItem "��", 7
    End If
    If ComBo.Text = "����" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "����", 0
    Com1.AddItem "ǧ��", 1
    Com1.AddItem "����������", 2
    Com1.AddItem "Ӣ��-��/����", 3
    Com1.AddItem "Ӣ��������λ/����", 4
    Com2.AddItem "����", 0
    Com2.AddItem "ǧ��", 1
    Com2.AddItem "����������", 2
    Com2.AddItem "Ӣ��-��/����", 3
    Com2.AddItem "Ӣ��������λ/����", 4
    End If
    If ComBo.Text = "����" Then
    Com1.Enabled = True
    Com2.Enabled = True
    Com1.AddItem "λ", 0
    Com1.AddItem "�ֽ�", 1
    Com1.AddItem "ǧλ", 2
    Com1.AddItem "ǧ�ֽ�", 3
    Com1.AddItem "��λ", 4
    Com1.AddItem "���ֽ�", 5
    Com1.AddItem "ǧ��λ", 6
    Com1.AddItem "ǧ���ֽ�", 7
    Com1.AddItem "ǧ��λ", 8
    Com1.AddItem "ǧ���ֽ�", 9
    Com1.AddItem "����λ", 10
    Com1.AddItem "�����ֽ�", 11
    Com2.AddItem "λ", 0
    Com2.AddItem "�ֽ�", 1
    Com2.AddItem "ǧλ", 2
    Com2.AddItem "ǧ�ֽ�", 3
    Com2.AddItem "��λ", 4
    Com2.AddItem "���ֽ�", 5
    Com2.AddItem "ǧ��λ", 6
    Com2.AddItem "ǧ���ֽ�", 7
    Com2.AddItem "ǧ��λ", 8
    Com2.AddItem "ǧ���ֽ�", 9
    Com2.AddItem "����λ", 10
    Com2.AddItem "�����ֽ�", 11
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
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ����Ӣ��" & vbCrLf & num2 & "  ����(Ӣ)   " & vbCrLf & num3 & "  Ʒ��(Ӣ)   "
                    If Format(num4 * 0.00000026667, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.00000026667, "0.00")) & "  ����Ӿ��"
                    Else
                        If Format(num4 * 0.002642, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.002642, "0.00")) & "  ��ԡ��"
                            Else
                                If Format(num4 * 4.226753, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 4.226753, "0.00")) & "  �����ȱ�"
                        End If
                    End If
            End If
            
            
            If ComBo.ListIndex = 1 Then
                num1 = Format(CDbl(L1) / cnu1 * 39.37008, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 3.28084, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 100, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  Ӣ��" & vbCrLf & num2 & "  Ӣ��" & vbCrLf & num3 & "  ����"
                    If Format(num4 * 0.01316, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.01316, "0.00")) & "  �ܴ�������ʽ�ͻ�"
                    Else
                        If Format(num4 * 5.351351, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 5.351351, "0.00")) & "  ֻ��"
                            Else
                                If Format(num4 * 28.545454, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 28.545454, "0.00")) & "  ��������"
                        End If
                    End If
            End If
            
            
            
            If ComBo.ListIndex = 2 Then
                num1 = Format(CDbl(L1) / cnu1 * 0.001, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 0.002205623, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 0.3527496, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ǧ��" & vbCrLf & num2 & "  ��" & vbCrLf & num3 & "  ��˾"
                 If Format(num4 * 0.0000000111, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0000000111, "0.00")) & "  ֻ��"
                    Else
                    If Format(num4 * 0.00000025, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.00000025, "0.00")) & "  ͷ����"
                    Else
                        If Format(num4 * 0.002312, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.002312, "0.00")) & "  ������"
                            Else
                                If Format(num4 * 500, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 500, "0.00")) & "  Ƭѩ��"
                        End If
                    End If
                    End If
            End If
            
            If ComBo.ListIndex = 3 Then
                num1 = Format(CDbl(L1) / cnu1, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 0.239005736137669, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 9.4816987913438E-04, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ����" & vbCrLf & num2 & "  ǧ��" & vbCrLf & num3 & "  BTU"

                    If Format(num4 * 0.0000009554, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0000009554, "0.00")) & "  ֻ����"
                    Else
                        If Format(num4 * 0.0000023, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0000023, "0.00")) & "  ֻ�㽶"
                            Else
                                If Format(num4 * 0.000111, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 0.000111, "0.00")) & "  �ڵ��"
                        End If
                    End If
             End If
                    

            If ComBo.ListIndex = 4 Then
                num1 = Format(CDbl(L1) / cnu1 * 0.0001, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 2.47105381467165E-04, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 1.19599004630108, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ����" & vbCrLf & num2 & "  ӢĶ" & vbCrLf & num3 & "  ƽ����"

                    If Format(num4 * 0.00001, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.00001, "0.00")) & "  ���Ǳ�"
                    Else
                        If Format(num4 * 16.58, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 16.58, "0.00")) & "  ��ֽ"
                            Else
                                If Format(num4 * 80, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 80, "0.00")) & "  ֻ��"
                        End If
                    End If
             End If
             
             
            If ComBo.ListIndex = 5 Then
                num1 = Format(CDbl(L1) / cnu1 * 1, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 2.9385836027035E-03, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 2.23713646532438, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ��/��" & vbCrLf & num2 & "  ���" & vbCrLf & num3 & "  Ӣ��/Сʱ"

                    If Format(num4 * 0.004068, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.004068, "0.00")) & "  ������ʽ�ɻ�"
                    Else
                        If Format(num4 * 0.0498, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0498, "0.00")) & "  ƥ��"
                            Else
                                If Format(num4 * 11.19, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 11.19, "0.00")) & "  ֻ�ڹ�"
                        End If
                    End If
             End If
             
            If ComBo.ListIndex = 6 Then
                num1 = Format(CDbl(L1) / cnu1 * 3600000, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 0.0416666666667, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 1.14077116130504E-04, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ��" & vbCrLf & num2 & "  ��" & vbCrLf & num3 & "  ��"
             End If
             
            If ComBo.ListIndex = 7 Then
                num1 = Format(CDbl(L1) / cnu1 * 3600000000#, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 1.34102208959503, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 56.8690192748062, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ����" & vbCrLf & num2 & "  ����" & vbCrLf & num3 & "  BTU/����"

                    If Format(num4 * 0.0003352666, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.0003352666, "0.00")) & "  ��������"
                    Else
                        If Format(num4 * 1.3413333, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 1.3413333, "0.00")) & "  ƥ��"
                            Else
                                If Format(num4 * 16.6455696203, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 16.6455696203, "0.00")) & "  ������"
                        End If
                    End If
             End If
             
            If ComBo.ListIndex = 8 Then
                num1 = Format(CDbl(L1) / cnu1 * 1024, "0.00")
                num2 = Format(CDbl(L1) / cnu1 * 134217728, "0.00")
                num3 = Format(CDbl(L1) / cnu1 * 1, "0.00")
                num4 = CStr(CSng(L1) / cnu1)
                If num1 = "0.00" Then num1 = "������1"
                If num2 = "0.00" Then num2 = "������1"
                If num3 = "0.00" Then num3 = "������1"
                Tip = "Լ����" & vbCrLf & num1 & "  ��λ" & vbCrLf & num2 & "  �ֽ�" & vbCrLf & num3 & "  ǧ��λ"

                    If Format(num4 * 0.026666667, "0.000") > 0.1 Then
                        Tip = Tip & vbCrLf & CStr(Format(num4 * 0.026666667, "0.00")) & "  ��DVD"
                    Else
                        If Format(num4 * 0.183, "0.000") > 0.1 Then
                            Tip = Tip & vbCrLf & CStr(Format(num4 * 0.183, "0.00")) & "  ��CD"
                            Else
                                If Format(num4 * 88.927943761, "0.000") > 0.1 Then Tip = Tip & vbCrLf & CStr(Format(num4 * 88.927943761, "0.00")) & "  ������"
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
'=====================================================================================================���涯��Ч��
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
'==================================move�¼�
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


'==================================down�¼�
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
'==================================up�¼�
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
