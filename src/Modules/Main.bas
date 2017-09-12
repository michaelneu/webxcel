Attribute VB_Name = "Main"
Public Sub Main()
    Dim server As HttpServer
    Set server = New HttpServer
    
    server.Serve 8080
End Sub
