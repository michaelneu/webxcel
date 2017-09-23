VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebControllerCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_controllers As Collection


Private Sub Class_Initialize()
    Set m_controllers = New Collection
End Sub


Public Sub AddController(controller As IWebController)
    m_controllers.Add controller
End Sub


Public Function GetMatchingController(requestUrl As String) As IWebController
    For Each controller In m_controllers
        If controller.MatchesUrl(requestUrl) Then
            Set GetMatchingController = controller
            Exit Function
        End If
    Next controller
    
    Set GetMatchingController = Nothing
End Function
