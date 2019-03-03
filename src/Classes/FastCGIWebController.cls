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
End Function

