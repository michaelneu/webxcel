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

Private Const UrlPrefix = "/workbook/"

Private m_wrm As WorksheetRelationshipMapper


Private Sub Class_Initialize()
    Set m_wrm = New WorksheetRelationshipMapper
End Sub


Private Function IWebController_MatchesUrl(requestUrl As String) As Boolean
    IWebController_MatchesUrl = requestUrl Like UrlPrefix & "*"
End Function


Private Function IWebController_ProcessRequest(request As HttpRequest) As HttpResponse
    Dim response As HttpResponse
    Set response = New HttpResponse
    
    Dim sheetName As String
    sheetName = StringExtensions.SubString(request.Url, Len(UrlPrefix))
    
    Dim slashIndex As Integer
    slashIndex = InStr(sheetName, "/")
    
    If slashIndex > 0 Then
        sheetName = Left(sheetName, slashIndex)
    End If
    
    Dim values As IJson
    Set values = m_wrm.All(sheetName)
    
    response.Body = values.ToJson
    response.Headers.AddHeader "Content-Type", "application/json"
    response.StatusCode = 200
    
    Set IWebController_ProcessRequest = response
End Function
