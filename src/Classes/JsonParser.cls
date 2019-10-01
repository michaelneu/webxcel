VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_null As JsonValue


Private Sub Class_Initialize()
    Set m_null = New JsonValue
    m_null.value = Null
End Sub


Private Function ConsumeObject(ByVal text As String, ByRef counter As Long) As String
    Dim buffer As String
    Dim openBrackets As Long
    Dim openBraces As Long
    
    Dim inString As Boolean
    inString = False
    
    For counter = counter To Len(text)
        Dim char As String
        char = StringExtensions.CharAt(text, counter)
        
        buffer = buffer & char
        
        If char = """" Then
            inString = Not inString
        End If
        
        If Not inString Then
            Select Case char
                Case "["
                    openBrackets = openBrackets + 1
                
                Case "]"
                    openBrackets = openBrackets - 1
                    
                Case "{"
                    openBraces = openBraces + 1
                    
                Case "}"
                    openBraces = openBraces - 1
            End Select
        End If
        
        If openBrackets = 0 And openBraces = 0 Then
            ConsumeObject = buffer
            Exit Function
        End If
    Next counter
    
    Err.Raise StatusCode.ErrorMalformedJson
End Function


Public Function ParseObject(obj As String) As JsonObject
    If Not StringExtensions.StartsWith(obj, "{") Or Not StringExtensions.EndsWith(obj, "}") Then
        Err.Raise StatusCode.ErrorMalformedJson
        Exit Function
    End If
    
    Set ParseObject = New JsonObject
    
    obj = StringExtensions.Substring(obj, 1)
    obj = Left(obj, Len(obj) - 1)
    
    Dim propertyName As String
    Dim value As IJson
    Dim valueBuffer As String
    Dim i As Long
    
    For i = 1 To Len(obj)
        Dim char As String
        char = StringExtensions.CharAt(obj, i)
        
        Select Case char
            Case """"
                If Len(propertyName) > 0 Then
                    Set value = ParseString(ConsumeString(obj, i))
                    
                    ParseObject.SetProperty propertyName, value
                    propertyName = ""
                Else
                    propertyName = ParseString(ConsumeString(obj, i)).value
                End If
                
            Case ":"
                If Len(propertyName) = 0 Then
                    Err.Raise StatusCode.ErrorMalformedJson
                    Exit Function
                End If
                
            Case "["
                Set value = ParseArray(ConsumeArray(obj, i))
                
                ParseObject.SetProperty propertyName, value
                propertyName = ""
                
            Case "{"
                Set value = ParseObject(ConsumeObject(obj, i))
                
                ParseObject.SetProperty propertyName, value
                propertyName = ""
            
            Case " ", vbTab, vbCr, vbLf, vbCrLf
                If Len(valueBuffer) > 0 And Len(propertyName) = 0 Then
                    Err.Raise StatusCode.ErrorMalformedJson
                    Exit Function
                End If
                
            Case ","
                If Len(valueBuffer) > 0 Then
                    If Len(propertyName) = 0 Then
                        Err.Raise StatusCode.ErrorMalformedJson
                        Exit Function
                    End If
                    
                    Set value = ParsePlainText(valueBuffer)
                    ParseObject.SetProperty propertyName, value
                    
                    propertyName = ""
                    valueBuffer = ""
                End If
            
            Case Else
                valueBuffer = valueBuffer & char
        End Select
    Next i
    
    If Len(valueBuffer) > 0 Then
        If Len(propertyName) = 0 Then
            Err.Raise StatusCode.ErrorMalformedJson
            Exit Function
        End If
        
        Set value = ParsePlainText(valueBuffer)
        ParseObject.SetProperty propertyName, value
    End If
End Function


Private Function ConsumeString(ByVal text As String, ByRef counter As Long) As String
    Dim buffer As String
    buffer = ""
    
    For counter = counter To Len(text)
        Dim char As String
        char = StringExtensions.CharAt(text, counter)
        
        buffer = buffer & char
        
        If char = """" And Len(buffer) > 1 Then
            ConsumeString = buffer
            Exit Function
        End If
    Next counter
    
    Err.Raise StatusCode.ErrorMalformedJson
End Function


Public Function ParseString(str As String) As JsonValue
    str = StringExtensions.Substring(str, 1)
    str = Left(str, Len(str) - 1)
    
    Set ParseString = New JsonValue
    ParseString.value = str
End Function


Private Function ConsumeArray(ByVal text As String, ByRef counter As Long) As String
    Dim buffer As String
    Dim openBrackets As Long
    Dim openBraces As Long
    
    Dim inString As Boolean
    inString = False
    
    For counter = counter To Len(text)
        Dim char As String
        char = StringExtensions.CharAt(text, counter)
        
        buffer = buffer & char
        
        If char = """" Then
            inString = Not inString
        End If
        
        If Not inString Then
            Select Case char
                Case "["
                    openBrackets = openBrackets + 1
                
                Case "]"
                    openBrackets = openBrackets - 1
                    
                Case "{"
                    openBraces = openBraces + 1
                    
                Case "}"
                    openBraces = openBraces - 1
            End Select
        End If
        
        If openBrackets = 0 And openBraces = 0 Then
            ConsumeArray = buffer
            Exit Function
        End If
    Next counter
    
    Err.Raise StatusCode.ErrorMalformedJson
End Function


Public Function ParseArray(arr As String) As JsonArray
    If Not StringExtensions.StartsWith(arr, "[") Or Not StringExtensions.EndsWith(arr, "]") Then
        Err.Raise StatusCode.ErrorMalformedJson
        Exit Function
    End If
    
    Set ParseArray = New JsonArray
    
    arr = StringExtensions.Substring(arr, 1)
    arr = Left(arr, Len(arr) - 1)
    
    Dim item As IJson
    Dim itemBuffer As String
    Dim i As Long
    
    For i = 1 To Len(arr)
        Dim char As String
        char = StringExtensions.CharAt(arr, i)
        
        Select Case char
            Case """"
                Set item = ParseString(ConsumeString(arr, i))
                ParseArray.AddItem item
                
            Case "["
                Set item = ParseArray(ConsumeArray(arr, i))
                ParseArray.AddItem item
                
            Case "{"
                Set item = ParseObject(ConsumeObject(arr, i))
                ParseArray.AddItem item
            
            Case " ", vbTab, vbCr, vbLf, vbCrLf
                If Len(valueBuffer) > 0 Then
                    Err.Raise StatusCode.ErrorMalformedJson
                    Exit Function
                End If
                
            Case ","
                If Len(itemBuffer) > 0 Then
                    Set item = ParsePlainText(itemBuffer)
                    ParseArray.AddItem item
                    
                    itemBuffer = ""
                End If
            
            Case Else
                itemBuffer = itemBuffer & char
        End Select
    Next i
    
    If Len(itemBuffer) > 0 Then
        Set item = ParsePlainText(itemBuffer)
        ParseArray.AddItem item
    End If
End Function


Private Function ParsePlainText(text As String) As JsonValue
    If LCase(text) = "null" Then
        Set ParsePlainText = m_null
    Else
        Dim textAsJsonValue As JsonValue
        Set textAsJsonValue = New JsonValue
        textAsJsonValue.value = text
        
        Set ParsePlainText = textAsJsonValue
    End If
End Function
