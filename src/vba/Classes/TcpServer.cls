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

Private m_wsa As WSADATA
Private m_serverSocket As Long
Private m_errors As TcpError


Private Sub Class_Initialize()
    Set m_errors = New TcpError

    Dim result As Long
    result = WSAStartup(257, m_wsa)
    
    If result <> 0 Then
        m_errors.RaiseErrorSetup
        Exit Sub
    End If
End Sub



Public Sub BindTo(ByVal port As Long, Optional ByVal backlog As Integer = 10)
    m_serverSocket = socket(AF_INET, SOCK_STREAM, 0)
    
    If m_serverSocket = INVALID_SOCKET Then
        m_errors.RaiseErrorCreateSocket
        Exit Sub
    End If
    
    Dim endpoint As sockaddr_in
    endpoint.sin_family = AF_INET
    endpoint.sin_addr.s_addr = INADDR_ANY
    endpoint.sin_port = htons(port)
    
    Dim bindResult As Long
    bindResult = bind(m_serverSocket, endpoint, 16)
    
    If bindResult <> 0 Then
        Dispose
        m_errors.RaiseErrorBind (port)
        Exit Sub
    End If
    
    Dim listenResult As Long
    listenResult = listen(m_serverSocket, backlog)
    
    If listenResult <> 0 Then
        Dispose
        m_errors.RaiseErrorListen
        Exit Sub
    End If
End Sub


Public Function AcceptTcpClient() As TcpClient
    Dim client As TcpClient
    Dim socket
    Dim socketAddress As sockaddr
    
    socket = accept(m_serverSocket, socketAddress, 16)
    
    If socket = -1 Then
        Dispose
        m_errors.RaiseErrorAccept
        Exit Function
    End If
    
    Set client = New TcpClient
    client.Initialize (socket)
    
    Set AcceptTcpClient = client
End Function


Public Sub Dispose()
    closesocket (m_serverSocket)
    WSACleanup
End Sub
