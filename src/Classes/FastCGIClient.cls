VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_clientSocket As TcpClient


Public Sub Initialize(clientSocket As TcpClient)
    m_clientSocket = clientSocket
End Sub


Public Function ReadTypeInfo() As FastCGITypeInfo
    Dim record As IFastCGIRecord
    Set record = New FastCGITypeInfo
    record.ReadFromTcpClient m_clientSocket
    Set ReadTypeInfo = record
End Function


Public Sub WriteBegin()
    Dim record As IFastCGIRecord
    Set record = New FastCGIBeginRequest
    record.WriteToTcpClient m_clientSocket
End Sub


Public Function ReadBegin() As FastCGIBeginRequest
    Dim record As IFastCGIRecord
    Set record = New FastCGIBeginRequest
    record.ReadFromTcpClient m_clientSocket
    Set ReadBegin = record
End Function


Public Sub WriteParams()
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


Public Function ReadParams() As FastCGIParams
    Dim record As IFastCGIRecord
    Set record = New FastCGIParams
    record.ReadFromTcpClient m_clientSocket
    Set ReadParams = record
End Function


Public Sub WriteInput(text As String)
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


Public Sub WriteOutput(text As String)
    Dim stdin As FastCGIStream
    Set stdin = New FastCGIStream
    stdin.StreamType = FastCGI.FASTCGI_TYPE_STDOUT
    stdin.Content = text

    Dim bytes As String
    Dim record As IFastCGIRecord
    Set record = stdin
    record.WriteToTcpClient m_clientSocket
End Sub


Public Sub ReadStream()
    Dim record As IFastCGIRecord
    Set record = New FastCGIStream
    record.ReadFromTcpClient m_clientSocket
End Sub


Public Sub WriteEnd()
    Dim record As IFastCGIRecord
    Set record = New FastCGIEndRequest
    record.WriteToTcpClient m_clientSocket
End Sub


Public Sub ReadEnd()
    Dim record As IFastCGIRecord
    Set record = New FastCGIEndRequest
    record.ReadFromTcpClient m_clientSocket
End Sub
