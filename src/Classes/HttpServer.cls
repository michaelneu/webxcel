VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_tcpServer As TcpServer


Private Sub Class_Initialize()
    Set m_tcpServer = New TcpServer
End Sub


Public Sub Serve(ByVal port As Long, Optional ByVal rootDirectory As String = ".")
    m_tcpServer.BindTo port, 100
    
    Do While True
        Dim client As TcpClient
        Set client = m_tcpServer.AcceptTcpClient()
        
        Dim requestText As String
        requestText = client.ReceiveString()
        
        Dim request As HttpRequest
        Set request = New HttpRequest
        request.Parse requestText
        
        Dim response As HttpResponse
        Set response = ProcessRequest(request)
        
        Dim responseText As String
        responseText = ""
        
        responseText = responseText & "HTTP/1.1 " & response.StatusCode & " Nobody Needs This Anyway" & vbCrLf
        responseText = responseText & "Server: Microsoft Excel/" & Application.Version & vbCrLf
        responseText = responseText & "Content-Length: " & Len(response.Body) & vbCrLf
        responseText = responseText & "Connection: close" & vbCrLf
        responseText = responseText & vbCrLf
        responseText = responseText & response.Body
        
        client.SendString responseText
        client.Dispose
        Exit Do
    Loop
    
    m_tcpServer.Dispose
End Sub


Private Function ProcessRequest(request As HttpRequest) As HttpResponse
    Dim response As HttpResponse
    Set response = New HttpResponse
    
    response.Body = "It works!"
    response.StatusCode = 200
    
    Set ProcessRequest = response
End Function
