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
    Dim header As FastCGIHeader
    Set header = New FastCGIHeader
    header.ProtocolVersion = 1
    header.RequestId = m_requestId
    header.MessageType = FastCGI.FASTCGI_TYPE_BEGIN_REQUEST
    header.ContentLength = 8
    header.PaddingLength = 0
    
    Dim record As IFastCGIRecord
    Dim bytes As String
    bytes = ""

    Set record = header
    bytes = bytes & record.ToBytes()
    
    Set record = New FastCGIBeginRequest
    bytes = bytes & record.ToBytes()
    
    m_clientSocket.SendString bytes
End Sub


Private Sub SendParams()
    Dim params As FastCGIParams
    Set params = New FastCGIParams
    
    params.Add "SERVER_SOFTWARE", "Microsoft Excel/" & Application.version
    params.Add "GATEWAY_INTERFACE", "CGI/1.1"

    Dim bytes As String
    Dim record As IFastCGIRecord
    Set record = params
    bytes = record.ToBytes()

    Dim header As FastCGIHeader
    Set header = New FastCGIHeader
    header.ProtocolVersion = 1
    header.RequestId = m_requestId
    header.MessageType = FastCGI.FASTCGI_TYPE_PARAMS
    header.ContentLength = Len(bytes)
    header.PaddingLength = 0

    Set record = header
    bytes = record.ToBytes() & bytes

    header.ContentLength = 0
    Set record = header
    bytes = bytes & record.ToBytes()

    m_clientSocket.SendString bytes
End Sub


Private Sub SendInput(text As String)
    Dim stdin As FastCGIStream
    Set stdin = New FastCGIStream
    stdin.Content = text

    Dim bytes As String
    Dim record As IFastCGIRecord
    Set record = stdin
    bytes = record.ToBytes()

    Dim header As FastCGIHeader
    Set header = New FastCGIHeader
    header.ProtocolVersion = 1
    header.RequestId = m_requestId
    header.MessageType = FastCGI.FASTCGI_TYPE_STDIN
    header.ContentLength = Len(bytes)
    header.PaddingLength = 0

    Set record = header
    bytes = record.ToBytes() & bytes

    If Len(text) > 0 Then
        header.ContentLength = 0
        Set record = header
        bytes = bytes & record.ToBytes()
    End If

    m_clientSocket.SendString bytes
End Sub
