Attribute VB_Name = "wsock32"
Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128

Public Const WSADESCRIPTION_LEN_ARRAY = WSADESCRIPTION_LEN + 1
Public Const WSASYS_STATUS_LEN_ARRAY = WSASYS_STATUS_LEN + 1

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN_ARRAY
    szSystemStatus As String * WSASYS_STATUS_LEN_ARRAY
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type

Public Const AF_INET = 2
Public Const SOCK_STREAM = 1
Public Const INADDR_ANY = 0

Public Type IN_ADDR
    s_addr As Long
End Type

Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero As String * 8
End Type

Public Type sockaddr
    sa_family As Integer
    sa_data As String * 14
End Type

Public Const INVALID_SOCKET = -1

Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal versionRequired As Long, wsa As WSADATA) As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function socket Lib "wsock32.dll" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long
Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function bind Lib "wsock32.dll" (ByVal socket As Long, name As sockaddr_in, ByVal nameLength As Integer) As Long
Public Declare Function listen Lib "wsock32.dll" (ByVal socket As Long, ByVal backlog As Integer) As Long
Public Declare Function accept Lib "wsock32.dll" (ByVal socket As Long, clientAddress As sockaddr, clientAddressLength As Integer) As Long
Public Declare Function send Lib "wsock32.dll" (ByVal socket As Long, buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare Function recv Lib "wsock32.dll" (ByVal socket As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
