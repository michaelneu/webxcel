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


Public Property Let StatusCode(ByVal value As Integer)
    m_statusCode = value
End Property


Public Property Get body() As String
    body = m_body
End Property


Public Property Let body(ByVal value As String)
    m_body = value
End Property


Public Property Get Headers() As HttpHeaderCollection
    Set Headers = m_headers
End Property


Public Function ToString() As String
    Headers.AddHeader "Server", "Microsoft Excel/" & Application.Version
    Headers.AddHeader "Content-Length", Len(body)
    Headers.AddHeader "Connection", "close"

    Dim text As String
    text = "HTTP/1.1 " & StatusCode & " " & HttpStatusMessages(StatusCode) & vbCrLf
    
    For Each header In Headers.GetEnumerator
        text = text & header.name & ": " & header.value & vbCrLf
    Next header
    
    text = text & vbCrLf
    text = text & body
    
    ToString = text
End Function


Public Sub Parse(ByVal response As String)
    Dim parts
    parts = Split(response, vbCrLf & vbCrLf, 2)
    
    If UBound(parts) <> 1 Then
        Err.Raise ErrorHttpRequestInvalidFormat
        Exit Sub
    End If
    
    ParseHeaders parts(0)
    m_body = Trim(parts(1))
End Sub


Private Sub ParseHeaders(ByVal Headers As String)
    Set m_headers = New HttpHeaderCollection

    Dim lines
    lines = Split(Headers, vbCrLf)
    
    For Each line In lines
        ParseHeader line
    Next
End Sub


Private Sub ParseHeader(ByVal line As String)
    line = Trim(line)

    Dim parts
    parts = Split(line, ":", 2)
    
    If parts(0) = "Status" Then
        StatusCode = Split(parts(0), " ", 1)
    Else
        m_headers.AddHeader Trim(parts(0)), Trim(parts(1))
    End If
End Sub
