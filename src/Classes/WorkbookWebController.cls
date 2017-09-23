VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookWebController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IWebController


Private Enum EntityRequestMode
    All
    ByKey
    Invalid
End Enum


Private Type EntityRequest
    Mode As EntityRequestMode
    Sheet As String
    Key As String
End Type


Private Const UrlPrefix = "/workbook/"

Private m_wrm As WorksheetRelationshipMapper


Private Sub Class_Initialize()
    Set m_wrm = New WorksheetRelationshipMapper
End Sub


Private Function IWebController_MatchesUrl(requestUrl As String) As Boolean
    IWebController_MatchesUrl = requestUrl Like UrlPrefix & "*"
End Function


Private Function IWebController_ProcessRequest(request As HttpRequest) As HttpResponse
    Dim requestHandled As Boolean
    requestHandled = False
    
    Dim response As HttpResponse
    Set response = New HttpResponse
    
    response.Headers.AddHeader "Content-Type", "application/json"
    
    Dim parsedRequest As EntityRequest
    parsedRequest = ParseRequestUrl(request.Url)
    
    If parsedRequest.Mode = All And request.RequestMethod = MethodGet Then
        Dim values As IJson
        Set values = m_wrm.All(parsedRequest.Sheet)
            
        response.Body = values.ToJson()
        response.StatusCode = 200
        
        requestHandled = True
    ElseIf parsedRequest.Mode = ByKey Then
        If request.RequestMethod = MethodGet Then
            Dim value As IJson
            Set value = m_wrm.Find(parsedRequest.Sheet, parsedRequest.Key)
            
            response.Body = value.ToJson()
            
            If TypeName(value) = "JsonValue" Then
                response.StatusCode = 404
            Else
                response.StatusCode = 200
            End If
            
            requestHandled = True
        End If
    End If
    
    If Not requestHandled Then
        response.StatusCode = 404
    End If
    
    Set IWebController_ProcessRequest = response
End Function


Private Function ParseRequestUrl(requestUrl As String) As EntityRequest
    requestUrl = StringExtensions.SubString(requestUrl, Len(UrlPrefix))
    
    Dim urlParts
    urlParts = Split(requestUrl, "/")
    
    Dim partCount As Integer
    partCount = UBound(urlParts) + 1
    
    If partCount > 2 Then
        ParseRequestUrl.Mode = Invalid
        Exit Function
    End If
    
    ParseRequestUrl.Sheet = urlParts(0)
    
    If partCount = 1 Then
        ParseRequestUrl.Mode = All
        Exit Function
    End If
    
    ParseRequestUrl.Key = Trim(urlParts(1))
    
    If Len(ParseRequestUrl.Key) = 0 Then
        ParseRequestUrl.Mode = All
    Else
        ParseRequestUrl.Mode = ByKey
    End If
End Function
