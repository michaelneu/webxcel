VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum HttpRequestMethod
    MethodGet
    MethodPost
    MethodPut
    MethodDelete
End Enum


Private m_requestMethod As HttpRequestMethod
Private m_url As String
Private m_headers As HttpHeaderCollection
Private m_body As String


Private Sub Class_Initialize()
    Set m_headers = New HttpHeaderCollection
End Sub


Public Property Get RequestMethod() As HttpRequestMethod
    RequestMethod = m_requestMethod
End Property


Public Property Get Url() As String
    Url = m_url
End Property


Public Property Get body() As String
    body = m_body
End Property


Public Property Get Headers() As HttpHeaderCollection
    Set Headers = m_headers
End Property


Public Sub Parse(ByVal request As String)
    request = Trim(request)
    
    Dim parts
    parts = Split(request, vbCrLf & vbCrLf, 2)
    
    If UBound(parts) <> 1 Then
        Err.Raise StatusCode.ErrorHttpRequestInvalidFormat
        Exit Sub
    End If
    
    ParseHeaders parts(0)
    m_body = Trim(parts(1))
End Sub


Private Sub ParseHeaders(text)
    text = Trim(text)
    
    Dim lines
    lines = Split(text, vbCrLf)
    
    Dim lineCount As Integer
    lineCount = UBound(lines)
    
    If lineCount = 0 Then
        Err.Raise StatusCode.ErrorHttpRequestInvalidFormat
    End If
    
    ParseProtocolLine lines(0)
    
    Dim i As Integer
    For i = 1 To lineCount
        ParseHeader lines(i)
    Next i
End Sub


Private Sub ParseProtocolLine(line)
    line = Trim(line)
    
    Dim upperLine As String
    upperLine = UCase(line)
    
    If Not upperLine Like "* HTTP/1.1" Then
        Err.Raise StatusCode.ErrorHttpRequestInvalidFormat
    End If
    
    Dim parts
    parts = Split(line, " ", 2)
    
    ParseRequestMethod parts(0)
    
    Dim urlPart As String
    urlPart = parts(1)
    
    Dim urlPartLength As Integer
    urlPartLength = Len(urlPart)
    
    Dim protocolSuffixLength As Integer
    protocolSuffixLength = Len(" HTTP/1.1")
    
    m_url = Left(urlPart, urlPartLength - protocolSuffixLength)
End Sub


Private Sub ParseHeader(line)
    line = Trim(line)
    
    Dim parts
    parts = Split(line, ":", 2)
    
    m_headers.AddHeader Trim(parts(0)), Trim(parts(1))
End Sub


Private Sub ParseRequestMethod(method)
    method = LCase(method)
    
    If method = "get" Then
        m_requestMethod = MethodGet
    ElseIf method = "post" Then
        m_requestMethod = MethodPost
    ElseIf method = "put" Then
        m_requestMethod = MethodPut
    ElseIf method = "delete" Then
        m_requestMethod = MethodDelete
    Else
        Err.Raise StatusCode.ErrorHttpRequestUnknownRequestMethod
    End If
End Sub
