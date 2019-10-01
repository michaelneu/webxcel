VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemWebController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IWebController

Private m_directory As String
Private m_indexFiles As Collection


Private Sub Class_Initialize()
    m_directory = ActiveWorkbook.Path
    
    Set m_indexFiles = New Collection
    m_indexFiles.Add "index.html"
    m_indexFiles.Add "index.htm"
End Sub


Public Property Get Directory() As String
    Directory = m_directory
End Property


Public Property Let Directory(val As String)
    val = StringExtensions.TrimRight(val, "/")
    val = StringExtensions.TrimRight(val, "\")
    
    m_directory = val
End Property


Public Property Get IndexFiles() As Collection
    Set IndexFiles = m_indexFiles
End Property


Private Function IWebController_MatchesUrl(requestUrl As String) As Boolean
    IWebController_MatchesUrl = True
End Function


Private Function IWebController_ProcessRequest(request As HttpRequest) As HttpResponse
    Dim response As HttpResponse
    Set response = New HttpResponse
    
    Dim filename As String
    filename = ResolveFile(request.Url)
    
    Dim file As FileInfo
    Set file = New FileInfo
    
    file.Initialize filename
    
    If file.Exists Then
        response.StatusCode = 200
        response.Headers.AddHeader "Content-Type", file.MimeType
        response.body = file.ReadString
    ElseIf m_indexFiles.count > 0 Then
        Dim foundIndexFile As Boolean
        foundIndexFile = False
        
        For Each indexFile In m_indexFiles
            Set file = New FileInfo
            file.Initialize PathJoin(filename, indexFile)

            If file.Exists Then
                response.StatusCode = 200
                response.Headers.AddHeader "Content-Type", file.MimeType
                response.body = file.ReadString
                
                foundIndexFile = True
                Exit For
            End If
        Next

        If foundIndexFile = False Then
            response.StatusCode = 404
        End If
    Else
        response.StatusCode = 404
    End If
    
    Set IWebController_ProcessRequest = response
End Function


Private Function ResolveFile(file As String) As String
    file = StringExtensions.TrimLeft(file, "/")
    file = StringExtensions.TrimLeft(file, "\")
    
    ResolveFile = m_directory & "\" & file
End Function
