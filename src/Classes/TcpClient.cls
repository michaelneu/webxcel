VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_clientSocket As Long


Public Sub Initialize(ByVal socket As Long)
    Dim timeout As Long
    timeout = 500
    
    Dim result As Long
    result = wsock32.setsockopt(socket, wsock32.SOL_SOCKET, wsock32.SO_RCVTIMEO, timeout, 4)
    
    m_clientSocket = socket
End Sub


Public Function SendString(ByVal message As String) As Long
    SendString = wsock32.send(m_clientSocket, ByVal message, Len(message), 0)
End Function


Public Function ReceiveString() As String
    Dim message As String
    Dim buffer As String * 1024
    Dim readBytes As Long
    
    message = ""
    
    Do
        buffer = ""
        readBytes = wsock32.recv(m_clientSocket, buffer, Len(buffer), 0)
        
        If readBytes > 0 Then
            message = message & Trim(buffer)
        End If
    Loop While readBytes > 0
    
    ReceiveString = Trim(message)
End Function


Public Sub Dispose()
    wsock32.closesocket (m_clientSocket)
End Sub
