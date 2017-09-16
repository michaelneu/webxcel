Attribute VB_Name = "StringExtensions"
Public Function TrimLeft(text As String, c As String) As String
    Dim textLength As Long
    textLength = Len(text)
    
    Do While textLength > 0
        Dim firstCharacter As String
        firstCharacter = Left(text, 1)
        
        If firstCharacter <> c Then
            Exit Do
        End If
        
        text = Right(text, textLength - 1)
        textLength = Len(text)
    Loop
    
    TrimLeft = text
End Function


Public Function TrimRight(text As String, c As String) As String
    Dim textLength As Long
    textLength = Len(text)
    
    Do While textLength > 0
        Dim lastCharacter As String
        lastCharacter = Left(text, 1)
        
        If lastCharacter <> c Then
            Exit Do
        End If
        
        text = Left(text, textLength - 1)
        textLength = Len(text)
    Loop
    
    TrimRight = text
End Function
