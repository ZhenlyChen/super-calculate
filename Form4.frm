VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form4 
   BackColor       =   &H00353535&
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":08CA
   ScaleHeight     =   3660
   ScaleWidth      =   6405
   StartUpPosition =   2  '��Ļ����
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   2295
      Left            =   7320
      TabIndex        =   3
      Top             =   960
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   4048
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
   Begin VB.Label up 
      BackStyle       =   0  'Transparent
      Caption         =   "V2.0beta"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Image min 
      Height          =   480
      Left            =   5040
      Picture         =   "Form4.frx":89CB
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Super����  �汾��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Image closes 
      Height          =   480
      Left            =   5760
      Picture         =   "Form4.frx":8AB5
      ToolTipText     =   "�ر�"
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwflags As Long, ByVal dwReserved As Long) As Long

 

Public Function GetConnectionString() As String
  Dim dwflags As Long
  Dim MyMsg As String
  '��ʾ����Ϣ�ַ���
     If InternetGetConnectedState(dwflags, 0&) Then
        If dwflags And INTERNET_CONNECTION_CONFIGURED Then
           MyMsg = MyMsg + "ϵͳ��������������" + vbCrLf
        End If
        If dwflags And INTERNET_CONNECTION_LAN Then
           MyMsg = MyMsg + "ϵͳͨ��·������Internet����"
        End If
        If dwflags And INTERNET_CONNECTION_PROXY Then
           MyMsg = MyMsg + "��ʹ����Proxy���������"
        Else
           MyMsg = MyMsg + "��"
        End If
        If dwflags And INTERNET_CONNECTION_MODEM Then
           MyMsg = MyMsg + "ϵͳʹ��Modem��Internet����"
        End If
        If dwflags And INTERNET_CONNECTION_OFFLINE Then
           MyMsg = MyMsg + "ϵͳ��ǰ��������״̬"
        End If
        If dwflags And INTERNET_CONNECTION_MODEM_BUSY Then
           MyMsg = MyMsg + "ϵͳͨADSL���ӵ�����"
        End If
        If dwflags And INTERNET_RAS_INSTALLED Then
           MyMsg = MyMsg + "ϵͳ��װ��Զ�̷��ʷ���"
        End If
     Else
        MyMsg = "ϵͳ��ǰδ��Internet����"
     End If
     GetConnectionString = MyMsg
End Function


Private Sub closes_Click()
Unload Me
End Sub

Private Sub closes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closes = Form2.closes(1)
min = Form2.min(0)
End Sub



Private Sub Label2_Click()
 If InternetCheckConnection("http://www.baidu.com/", 1, 0) = 0 Then
     MsgBox "���粻����,�޷����ӷ�����" + vbCrLf + "��ϸ��Ϣ��" + GetConnectionString
    Else
Label2.Enabled = False
Label2 = "������..."
Dim StringA As String
Dim StringB As String
Web.Navigate ("http://zhen.whostii.com/client/up.html")
DoEvents
While Web.Busy
    DoEvents
Wend
Web.Document.body.focus
Web.Document.execCommand "SelectAll"
Web.Document.execCommand "copy"
StringA = Clipboard.GetText
'StringA = WebDaima(Web, "All")
Web.Navigate ("http://zhen.whostii.com/client/uptip.html")
DoEvents
While Web.Busy
    DoEvents
Wend
Web.Document.body.focus
Web.Document.execCommand "SelectAll"
Web.Document.execCommand "copy"
StringB = Clipboard.GetText
'StringB = WebDaima(Web, "All")

StringA = Left(StringA, 4)
'StringB = findStr(StringB, "<PRE>", "</PRE>")
If up <> StringA Then
msg = MsgBox("�����°汾" & StringA & "���Ƿ�����" + vbCrLf + StringB, vbYesNo, "��ʾ")
If msg = vbYes Then
Label2 = "������...."
R = URLDownloadToFile(0, "http://zhen.whostii.com/client/new" & StringA, App.Path & "\" & "Super����" & StringA & ".exe", 0, 0)
MsgBox "�������"
Open App.Path & "\old.tmp" For Output As #1
Print #1, App.EXEName & ".exe"
Close
Shell (App.Path & "\" & "Super����" & StringA & ".exe"), 1
End
End If
Label2 = "������"
Label2.Enabled = True
Else
MsgBox ("��ϲ�㣬��İ汾�����µ�")
Label2 = "������"
End If
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF00&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFF&
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFFFF
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
End Sub
Public Function WebDaima(WebBrowser, BuFen)
  Select Case BuFen
    Case "Body"
      WebDaima = WebBrowser.Document.body.innerhtml
    Case "All"
      WebDaima = WebBrowser.Document.documentelement.outerhtml
    Case Else
      WebDaima = WebBrowser.Document.documentelement.outerhtml
  End Select
End Function
Private Function findStr(str1 As String, str2 As String, str3 As String)
    Dim intStart, intEnd As Integer
    If InStr(1, str1, str2) = 0 Or InStr(1, str1, str3) = 0 Then Exit Function
    intStart = InStr(1, str1, str2) + Len(str2)
    intEnd = InStr(1, str1, str3)
    findStr = Mid(str1, intStart, intEnd - intStart)
End Function
