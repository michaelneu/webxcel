VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_wsa As wsock32.WSADATA
Private m_serverSocket As Long


Private Sub Class_Initialize()
    Dim result As Long
    result = wsock32.WSAStartup(257, m_wsa)
    
    If result <> 0 Then
        Err.Raise StatusCode.ErrorSocketSetup
        Exit Sub
    End If
End Sub



Public Sub BindTo(ByVal port As Long, Optional ByVal backlog As Integer = 10)
    m_serverSocket = wsock32.socket(AF_INET, SOCK_STREAM, 0)
    
    If m_serverSocket = wsock32.INVALID_SOCKET Then
        Err.Raise StatusCode.ErrorSocketCreation
        Exit Sub
    End If
    
    Dim endpoint As wsock32.sockaddr_in
    endpoint.sin_family = wsock32.AF_INET
    endpoint.sin_addr.s_addr = wsock32.INADDR_ANY
    endpoint.sin_port = wsock32.htons(port)
    
    Dim bindResult As Long
    bindResult = wsock32.bind(m_serverSocket, endpoint, 16)
    
    If bindResult <> 0 Then
        Dispose
        Err.Raise StatusCode.ErrorSocketBind
        Exit Sub
    End If
    
    Dim listenResult As Long
    listenResult = wsock32.listen(m_serverSocket, backlog)
    
    If listenResult <> 0 Then
        Dispose
        Err.Raise StatusCode.ErrorSocketStartListening
        Exit Sub
    End If
End Sub


Public Function AcceptTcpClient() As TcpClient
    Dim client As TcpClient
    Dim socket
    Dim socketAddress As wsock32.sockaddr
    
    socket = wsock32.accept(m_serverSocket, socketAddress, 16)
    
    If socket = -1 Then
        Dispose
        Err.Raise StatusCode.ErrorSocketAcceptClient
        Exit Function
    End If
    
    Set client = New TcpClient
    client.Initialize (socket)
    
    Set AcceptTcpClient = client
End Function


Public Sub Dispose()
    wsock32.closesocket (m_serverSocket)
    wsock32.WSACleanup
End Sub
