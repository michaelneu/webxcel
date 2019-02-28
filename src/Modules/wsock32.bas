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

Public Const FD_SETSIZE = 64

Public Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Long
End Type

Public Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Public Type sockaddr
    sa_family As Integer
    sa_data As String * 14
End Type

Public Const INVALID_SOCKET = -1

Public Const SOL_SOCKET = 65535
Public Const SO_RCVTIMEO = &H1006

Public Declare PtrSafe Function WSAStartup Lib "wsock32.dll" (ByVal versionRequired As Long, wsa As WSADATA) As Long
Public Declare PtrSafe Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare PtrSafe Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare PtrSafe Function socket Lib "wsock32.dll" (ByVal addressFamily As Long, ByVal socketType As Long, ByVal protocol As Long) As Long
Public Declare PtrSafe Function connect Lib "wsock32.dll" (ByVal s As Long, ByRef address As sockaddr_in, ByVal namelen As Long) As Long
Public Declare PtrSafe Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Public Declare PtrSafe Function bind Lib "wsock32.dll" (ByVal socket As Long, name As sockaddr_in, ByVal nameLength As Integer) As Long
Public Declare PtrSafe Function listen Lib "wsock32.dll" (ByVal socket As Long, ByVal backlog As Integer) As Long
Public Declare PtrSafe Function select_ Lib "wsock32.dll" Alias "select" (ByVal nfds As Integer, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Integer
Public Declare PtrSafe Function accept Lib "wsock32.dll" (ByVal socket As Long, clientAddress As sockaddr, clientAddressLength As Integer) As Long
Public Declare PtrSafe Function setsockopt Lib "wsock32.dll" (ByVal socket As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Integer) As Long
Public Declare PtrSafe Function send Lib "wsock32.dll" (ByVal socket As Long, buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function recv Lib "wsock32.dll" (ByVal socket As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function inet_addr Lib "wsock32.dll" (ByVal hostname As String) As Long
Public Declare PtrSafe Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long


Public Sub FD_ZERO_MACRO(ByRef s As fd_set)
    s.fd_count = 0
End Sub


Public Sub FD_SET_MACRO(ByVal fd As Long, ByRef s As fd_set)
    Dim i As Integer
    i = 0
    
    Do While i < s.fd_count
        If s.fd_array(i) = fd Then
            Exit Do
        End If
        
        i = i + 1
    Loop
    
    If i = s.fd_count Then
        If s.fd_count < FD_SETSIZE Then
            s.fd_array(i) = fd
            s.fd_count = s.fd_count + 1
        End If
    End If
End Sub
