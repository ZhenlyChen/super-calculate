Attribute VB_Name = "Module3"
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
'��ϵͳ��������״̬
Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20

'��װ��RAS����
Public Const INTERNET_RAS_INSTALLED As Long = &H10

'ʹ��Modem��Internet����
Public Const INTERNET_CONNECTION_MODEM As Long = &H1

'ʹ��Proxy��Internet����
Public Const INTERNET_CONNECTION_PROXY As Long = &H4

'δʹ��
Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8

'ʹ��LAN��Internet����
Public Const INTERNET_CONNECTION_LAN As Long = &H2

'�����кϷ���Internet�������ã�����ǰ���ܴ�������״̬������״̬��
Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
