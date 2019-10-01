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


Public host As String
Public port As Long
Public Extension As String


Private Function IWebController_MatchesUrl(requestUrl As String) As Boolean
    If requestUrl Like Extension Then
        IWebController_MatchesUrl = True
    Else
        IWebController_MatchesUrl = False
    End If
End Function


Private Function IWebController_ProcessRequest(request As HttpRequest) As HttpResponse
    Dim clientSocket As TcpClient
    Set clientSocket = New TcpClient
    clientSocket.ConnectTo host, port

    Dim fcgiClient As FastCGIClient
    Set fcgiClient = New FastCGIClient
    fcgiClient.Initialize clientSocket
    
    fcgiClient.WriteBegin
    
    Dim params As FastCGIParams
    Set params = New FastCGIParams

    params.Add "REQUEST_METHOD", request.RequestMethod
    params.Add "SERVER_PROTOCOL", "http"
    params.Add "SERVER_NAME", "webxcel.local"
    params.Add "SERVER_SOFTWARE", "Microsoft Excel/" & Application.Version
    params.Add "GATEWAY_INTERFACE", "CGI/1.1"
    params.Add "SCRIPT_NAME", request.Url
    params.Add "SCRIPT_FILENAME", request.Url
    params.Add "DOCUMENT_URI", request.Url
    
    For Each header In request.Headers.GetEnumerator
        params.Add ConvertHeaderNameToParamName(header.name), header.value
    Next
    
    fcgiClient.WriteParams params
    Set params = New FastCGIParams
    fcgiClient.WriteParams params

    If Len(request.body) > 0 Then
        Dim stdin As String
        stdin = request.body

        Do While Len(stdin) > 65535
            Dim chunk As String
            chunk = StringExtensions.Substring(stdin, 0, 65535)
            fcgiClient.WriteInput chunk
            stdin = StringExtensions.Substring(stdin, 65535, 0)
        Loop
    End If

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
    
    If response.StatusCode = 0 Then
        response.StatusCode = 200
    End If

    Set IWebController_ProcessRequest = response
End Function


Private Function ConvertHeaderNameToParamName(ByVal headerName As String) As String
    ConvertHeaderNameToParamName = UCase(Replace(headerName, "-", "_"))
End Function
