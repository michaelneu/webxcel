Attribute VB_Name = "Main"
Public Sub Main()
    Dim server As HttpServer
    Set server = New HttpServer
    
    server.Controllers.AddController New WorkbookWebController
    server.Controllers.AddController New FileSystemWebController
    
    server.Serve 8080
End Sub
