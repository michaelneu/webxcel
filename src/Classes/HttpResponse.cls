VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_statusCode As Integer
Private m_body As String
Private m_headers As HttpHeaderCollection


Private Sub Class_Initialize()
    Set m_headers = New HttpHeaderCollection
End Sub


Public Property Get StatusCode() As Integer
    StatusCode = m_statusCode
End Property


Public Property Let StatusCode(ByVal Value As Integer)
    m_statusCode = Value
End Property


Public Property Get Body() As String
    Body = m_body
End Property


Public Property Let Body(ByVal Value As String)
    m_body = Value
End Property


Public Property Get Headers() As HttpHeaderCollection
    Set Headers = m_headers
End Property

Public Function ToString() As String
    Headers.AddHeader "Server", "Microsoft Excel/" & Application.Version
    Headers.AddHeader "Content-Length", Len(Body)
    Headers.AddHeader "Connection", "close"

    Dim text As String
    text = "HTTP/1.1 " & StatusCode & " " & HttpStatusMessages(StatusCode) & vbCrLf
    
    For Each header In Headers.GetEnumerator
        text = text & header.name & ": " & header.Value & vbCrLf
    Next header
    
    text = text & vbCrLf
    text = text & Body
    
    ToString = text
End Function

Private Function HeadersToString() As String

End Function
