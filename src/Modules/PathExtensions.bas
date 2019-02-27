Attribute VB_Name = "PathExtensions"
Public Function PathJoin(ParamArray pathParts() As Variant) As String
    If IsEmpty(pathParts) Then
        Exit Function
    End If

    Dim partIndex As Integer
    Dim pathPart As String
    Dim joinedPath As String
    joinedPath = ""

    For partIndex = LBound(pathParts) To UBound(pathParts) - 1
        pathPart = pathParts(partIndex)
        pathPart = TrimRight(pathPart, "/")
        pathPart = TrimRight(pathPart, "\")
        
        joinedPath = joinedPath & pathPart & "/"
    Next

    pathPart = pathParts(UBound(pathParts))

    If UBound(pathParts) > 1 Then
        pathPart = TrimLeft(pathPart, "/")
        pathPart = TrimLeft(pathPart, "\")
    End If

    joinedPath = joinedPath & pathPart
    PathJoin = joinedPath
End Function
