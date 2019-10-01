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
    key As String
End Type


Private Const UrlPrefix = "/workbook/"

Private m_wrm As WorksheetRelationshipMapper
Private m_jsonParser As JsonParser


Private Sub Class_Initialize()
    Set m_wrm = New WorksheetRelationshipMapper
    Set m_jsonParser = New JsonParser
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
    
    Dim isJsonRequest As Boolean
    isJsonRequest = RequestContainsJsonData(request)
    
    If parsedRequest.Mode = All Then
        If request.RequestMethod = MethodGet Then
            Dim values As IJson
            Set values = m_wrm.All(parsedRequest.Sheet)
                
            response.StatusCode = 200
            response.body = values.ToJson()
            
            requestHandled = True
        ElseIf request.RequestMethod = MethodPost And isJsonRequest Then
            Dim postBody As JsonObject
            Set postBody = m_jsonParser.ParseObject(request.body)
            
            m_wrm.Insert parsedRequest.Sheet, postBody
            
            Dim responseValue As IJson
            Set responseValue = postBody
            
            Dim insertedId As JsonValue
            Set insertedId = postBody.GetProperty(m_wrm.PrimaryKeyName(parsedRequest.Sheet))
            
            response.Headers.AddHeader "Location", UrlPrefix & parsedRequest.Sheet & "/" & insertedId.value
            response.StatusCode = 201
            response.body = responseValue.ToJson()
            
            requestHandled = True
        End If
    ElseIf parsedRequest.Mode = ByKey Then
        If request.RequestMethod = MethodGet Then
            Dim getValue As IJson
            Set getValue = m_wrm.Find(parsedRequest.Sheet, parsedRequest.key)
            
            response.body = getValue.ToJson()
            
            ' this is a null value check in disguise
            If TypeName(getValue) = "JsonValue" Then
                response.StatusCode = 404
            Else ' type is JsonObject
                response.StatusCode = 200
            End If
            
            requestHandled = True
        ElseIf request.RequestMethod = MethodDelete Then
            Dim deleteValue As IJson
            Set deleteValue = m_wrm.Delete(parsedRequest.Sheet, parsedRequest.key)
            
            If deleteValue Is Nothing Then
                response.StatusCode = 404
            Else
                response.StatusCode = 200
                response.Headers.AddHeader "Content-Type", "application/json"
                response.body = deleteValue.ToJson()
            End If
            
            requestHandled = True
        ElseIf request.RequestMethod = MethodPut And isJsonRequest Then
            Dim putBody As JsonObject
            Set putBody = m_jsonParser.ParseObject(request.body)
            
            If m_wrm.Update(parsedRequest.Sheet, parsedRequest.key, putBody) Then
                response.StatusCode = 204
            Else
                response.StatusCode = 404
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
    requestUrl = StringExtensions.Substring(requestUrl, Len(UrlPrefix))
    
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
    
    ParseRequestUrl.key = Trim(urlParts(1))
    
    If Len(ParseRequestUrl.key) = 0 Then
        ParseRequestUrl.Mode = All
    Else
        ParseRequestUrl.Mode = ByKey
    End If
End Function


Private Function RequestContainsJsonData(request As HttpRequest) As Boolean
    RequestContainsJsonData = False
    
    For Each header In request.Headers.GetEnumerator
        Dim n As String
        n = LCase(header.name)
        
        Dim v As String
        v = LCase(header.value)
        
        If LCase(header.name) = "content-type" And (LCase(header.value) = "application/json" Or LCase(header.value) = "application/json;charset=utf-8") Then
            RequestContainsJsonData = True
            Exit Function
        End If
    Next header
End Function
