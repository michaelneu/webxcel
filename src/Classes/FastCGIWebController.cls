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

Private m_requestId As Integer
Public Host As String
Public Port As Long
Private m_clientSocket As TcpClient


Private Sub Class_Initialize()
    m_requestId = 0
End Sub


Private Function IWebController_MatchesUrl(requestUrl As String) As Boolean
    IWebController_MatchesUrl = True
End Function


Private Function IWebController_ProcessRequest(request As HttpRequest) As HttpResponse
    m_requestId = m_requestId + 1
End Function


Private Sub SendBegin()
    Dim record As IFastCGIRecord
    Set record = New FastCGIBeginRequest

    record.WriteToTcpClient client
End Sub


Private Sub SendParams()
    Dim params As FastCGIParams
    Set params = New FastCGIParams

    params.Add "SERVER_SOFTWARE", "Microsoft Excel/" & Application.version
    params.Add "GATEWAY_INTERFACE", "CGI/1.1"

    Dim record As IFastCGIRecord
    Set record = params
    record.WriteToTcpClient client

    Set record = New FastCGIParams
    record.WriteToTcpClient client
End Sub


Private Sub SendInput(text As String)
    Dim stdin As FastCGIStream
    Set stdin = New FastCGIStream
    stdin.StreamType = FastCGI.FASTCGI_TYPE_STDIN
    stdin.Content = text

    Dim bytes As String
    Dim record As IFastCGIRecord
    Set record = stdin
    record.WriteToTcpClient client

    If Len(text) > 0 Then
        stdin.Content = ""
        Set record = stdin
        record.WriteToTcpClient client
    End If
End Sub
