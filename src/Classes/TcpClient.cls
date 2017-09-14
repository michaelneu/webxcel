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
    m_clientSocket = socket
End Sub


Public Sub SendString(ByVal Message As String)
    Dim result As Long
    
    result = wsock32.send(m_clientSocket, ByVal Message, Len(Message), 0)
End Sub


Public Function ReceiveString() As String
    Dim Message As String
    Dim buffer As String * 1024
    Dim readBytes As Long
    
    Message = ""
    
    Do
        readBytes = wsock32.recv(m_clientSocket, buffer, Len(buffer), 0)
        
        If readBytes > 0 Then
            Message = Message & buffer
        End If
    Loop While readBytes = Len(buffer)
    
    ReceiveString = Message
End Function


Public Sub Dispose()
    wsock32.closesocket (m_clientSocket)
End Sub
