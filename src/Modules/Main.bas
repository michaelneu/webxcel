Attribute VB_Name = "Main"
Public Sub Main()
    Dim server As HttpServer
    Set server = New HttpServer

    Dim php As FastCGIWebController
    Set php = New FastCGIWebController

    php.host = "localhost"
    php.port = 9000
    php.Extension = "*.php"

    server.Controllers.AddController php
    server.Controllers.AddController New WorkbookWebController
    server.Controllers.AddController New FileSystemWebController

    server.Serve 8080
End Sub
