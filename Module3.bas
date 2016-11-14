Attribute VB_Name = "Module3"
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
'本系统处于离线状态
Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20

'安装了RAS服务
Public Const INTERNET_RAS_INSTALLED As Long = &H10

'使用Modem与Internet相连
Public Const INTERNET_CONNECTION_MODEM As Long = &H1

'使用Proxy与Internet相连
Public Const INTERNET_CONNECTION_PROXY As Long = &H4

'未使用
Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8

'使用LAN与Internet相连
Public Const INTERNET_CONNECTION_LAN As Long = &H2

'主机有合法的Internet连接配置（但当前可能处于连接状态或离线状态）
Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
