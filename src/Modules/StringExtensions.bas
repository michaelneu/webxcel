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
        lastCharacter = Right(text, 1)
        
        If lastCharacter <> c Then
            Exit Do
        End If
        
        text = Left(text, textLength - 1)
        textLength = Len(text)
    Loop
    
    TrimRight = text
End Function


Public Function Substring(ByVal text As String, startIndex As Integer, Optional length As Variant) As String
    If startIndex > Len(text) Then
        startIndex = Len(text)
    End If

    If IsMissing(length) Or TypeName(length) <> "Integer" Or length > Len(text) Then
        length = Len(text) - startIndex
    End If

    Substring = Left(Right(text, Len(text) - startIndex), length)
End Function


Public Function StartsWith(ByVal text As String, ByVal startText As String) As Boolean
    StartsWith = InStr(text, startText) = 1
End Function


Public Function EndsWith(ByVal text As String, ByVal endText As String) As Boolean
    EndsWith = Right(text, Len(endText)) = endText
End Function


Public Function CharAt(ByVal text As String, ByVal index As Integer) As String
    CharAt = Mid(text, index, 1)
End Function


Public Function Repeat(ByVal text As String, ByVal count As Long) As String
    Repeat = ""
    
    Dim i As Long
    For i = 1 To count
        Repeat = Repeat & text
    Next
End Function
