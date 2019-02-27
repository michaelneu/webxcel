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
Private m_controllers As WebControllerCollection


Private Sub Class_Initialize()
    InitializeHttpStatusDictionary

    Set m_tcpServer = New TcpServer
    Set m_controllers = New WebControllerCollection
End Sub


Public Property Get Controllers() As WebControllerCollection
    Set Controllers = m_controllers
End Property


Public Sub Serve(ByVal port As Long)
    Dim lockfile As FileInfo
    Set lockfile = New FileInfo
    
    lockfile.Initialize ActiveWorkbook.FullName & ".lock"
    lockfile.Create
    
    m_tcpServer.BindTo port, 100
    
    Do While True
        Dim client As TcpClient
        Set client = m_tcpServer.AcceptTcpClient(1000)
        
        If Not client Is Nothing Then
            Dim requestText As String
            requestText = client.ReceiveString()
            
            If requestText <> "" Then
                Dim request As HttpRequest
                Set request = New HttpRequest
                request.Parse requestText
                
                Dim response As HttpResponse
                Set response = ProcessRequest(request)
                
                Dim responseText As String
                responseText = response.ToString()
                
                client.SendString responseText
            End If
            
            client.Dispose
        End If
        
        If Not lockfile.Exists Then
            Exit Do
        End If
        
        DoEvents
    Loop
    
    m_tcpServer.Dispose
End Sub


Private Function ProcessRequest(request As HttpRequest) As HttpResponse
    Dim controller As IWebController
    Set controller = Controllers.GetMatchingController(request.Url)
    
    If Not controller Is Nothing Then
        Set ProcessRequest = controller.ProcessRequest(request)
    Else
        Dim response As HttpResponse
        Set response = New HttpResponse
        
        response.StatusCode = 500
        
        Set ProcessRequest = response
    End If
End Function
