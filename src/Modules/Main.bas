Attribute VB_Name = "Main"
Public Sub Main()
    Dim server As HttpServer
    Set server = New HttpServer
    
    Dim fileController As FileSystemWebController
    Set fileController = New FileSystemWebController
    
    server.Controllers.AddController fileController
    server.Serve 8080
End Sub
