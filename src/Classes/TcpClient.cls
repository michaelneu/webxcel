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


Public Sub SendString(ByVal message As String)
    Dim result As Long
    
    result = wsock32.send(m_clientSocket, ByVal message, Len(message), 0)
End Sub


Public Function ReceiveString() As String
    Dim message As String
    Dim buffer As String * 1024
    Dim readBytes As Long
    
    message = ""
    
    Do
        readBytes = wsock32.recv(m_clientSocket, buffer, Len(buffer), 0)
        
        If readBytes > 0 Then
            message = message & buffer
        End If
    Loop While readBytes = Len(buffer)
    
    ReceiveString = TrimNullCharacters(message)
End Function


Private Function TrimNullCharacters(ByVal text As String) As String
    Dim nullCharacterPosition As Long
    nullCharacterPosition = InStr(text, Chr(0))

    If nullCharacterPosition = 0 Then
        text = ""
    ElseIf nullCharacterPosition > 0 Then
        text = Left(text, nullCharacterPosition - 1)
    End If

    TrimNullCharacters = text
End Function


Public Sub Dispose()
    wsock32.closesocket (m_clientSocket)
End Sub
