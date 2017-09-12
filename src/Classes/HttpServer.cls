VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum HttpRequestMethod
    MethodGet
    MethodPost
    MethodPut
    MethodDelete
End Enum


Private Type HttpRequest
    Method As HttpRequestMethod
    Url As String
End Type


Private Type HttpResponse
    StatusCode As Integer
    Body As String
End Type


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
        request = ParseRequest(requestText)
        
        Dim response As HttpResponse
        response = ProcessRequest(request)
        
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


Private Function ParseRequest(ByVal request As String) As HttpRequest
    Dim requestAsLower As String
    requestAsLower = LCase(request)
    
    ParseRequest.Method = ParseMethod(requestAsLower)
End Function


Private Function ParseMethod(ByVal requestAsLower As String) As HttpRequestMethod
    If requestAsLower Like "get *" Then
        ParseMethod = MethodGet
    ElseIf requestAsLower Like "post *" Then
        ParseMethod = MethodPost
    ElseIf requestAsLower Like "put *" Then
        ParseMethod = MethodPut
    ElseIf requestAsLower Like "delete *" Then
        ParseMethod = MethodDelete
    End If
End Function


Private Function ProcessRequest(request As HttpRequest) As HttpResponse
    ProcessRequest.Body = "It works!"
    ProcessRequest.StatusCode = 200
End Function
