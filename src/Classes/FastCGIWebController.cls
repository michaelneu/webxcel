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


Private Sub WriteBegin()
    Dim record As IFastCGIRecord
    Set record = New FastCGIBeginRequest
    record.WriteToTcpClient m_clientSocket
End Sub


Private Function ReadBegin() As FastCGIBeginRequest
    Dim record As IFastCGIRecord
    Set record = New FastCGIBeginRequest
    record.ReadFromTcpClient m_clientSocket
    Set ReadBegin = record
End Function


Private Sub WriteParams()
    Dim params As FastCGIParams
    Set params = New FastCGIParams

    params.Add "SERVER_SOFTWARE", "Microsoft Excel/" & Application.version
    params.Add "GATEWAY_INTERFACE", "CGI/1.1"

    Dim record As IFastCGIRecord
    Set record = params
    record.WriteToTcpClient m_clientSocket

    Set record = New FastCGIParams
    record.WriteToTcpClient m_clientSocket
End Sub


Private Function ReadParams() As FastCGIParams
    Dim record As IFastCGIRecord
    Set record = New FastCGIParams
    record.ReadFromTcpClient m_clientSocket
    Set ReadParams = record
End Function


Private Sub WriteInput(text As String)
    Dim stdin As FastCGIStream
    Set stdin = New FastCGIStream
    stdin.StreamType = FastCGI.FASTCGI_TYPE_STDIN
    stdin.Content = text

    Dim bytes As String
    Dim record As IFastCGIRecord
    Set record = stdin
    record.WriteToTcpClient m_clientSocket

    If Len(text) > 0 Then
        stdin.Content = ""
        Set record = stdin
        record.WriteToTcpClient m_clientSocket
    End If
End Sub


Private Sub WriteOutput(text As String)
    Dim stdin As FastCGIStream
    Set stdin = New FastCGIStream
    stdin.StreamType = FastCGI.FASTCGI_TYPE_STDOUT
    stdin.Content = text

    Dim bytes As String
    Dim record As IFastCGIRecord
    Set record = stdin
    record.WriteToTcpClient m_clientSocket
End Sub


Private Sub ReadStream()
    Dim record As IFastCGIRecord
    Set record = New FastCGIStream
    record.ReadFromTcpClient m_clientSocket
End Sub


Private Sub WriteEnd()
    Dim record As IFastCGIRecord
    Set record = New FastCGIEndRequest
    record.WriteToTcpClient m_clientSocket
End Sub


Private Sub ReadEnd()
    Dim record As IFastCGIRecord
    Set record = New FastCGIEndRequest
    record.ReadFromTcpClient m_clientSocket
End Sub

