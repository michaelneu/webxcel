VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIWebController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IWebController


Public Host As String
Public Port As Long


Private Function IWebController_MatchesUrl(requestUrl As String) As Boolean
    IWebController_MatchesUrl = True
End Function


Private Function IWebController_ProcessRequest(request As HttpRequest) As HttpResponse
    Dim clientSocket As TcpClient
    Set clientSocket = New TcpClient
    clientSocket.ConnectTo Host, Port

    Dim fcgiClient As FastCGIClient
    Set fcgiClient = New FastCGIClient
    fcgiClient.Initialize clientSocket
    
    fcgiClient.WriteBegin
    fcgiClient.WriteParams
    fcgiClient.WriteInput ""
    
    Dim typeInfo As FastCGITypeInfo
    Dim body As String
    body = ""
    
    Do While True
        Set typeInfo = fcgiClient.ReadTypeInfo()
        
        If typeInfo.MessageType = FastCGI.FASTCGI_TYPE_STDOUT Then
            Dim stdout As FastCGIStream
            Set stdout = fcgiClient.ReadStream()
            body = body & stdout.Content
        ElseIf typeInfo.MessageType = FastCGI.FASTCGI_TYPE_END_REQUEST Then
            Dim endRequest As FastCGIEndRequest
            Set endRequest = fcgiClient.ReadEnd()

            Exit Do
        End If
    Loop

    clientSocket.Dispose

    Dim response As HttpResponse
    Set response = New HttpResponse
    response.Parse body
    response.StatusCode = 200

    Set IWebController_ProcessRequest = response
End Function
