Attribute VB_Name = "StringExtensions"
' Trims the given character from the given text starting from the left.
' Arguments:
' - text The text to trim
' - c The character to trim
Public Function TrimLeft(ByVal text As String, c As String) As String
    Dim textLength As Long
    textLength = Len(text)
    
    Dim firstCharacter As String
    
    Do While textLength > 0
        firstCharacter = Left(text, 1)
        
        If firstCharacter <> c Then
            Exit Do
        End If
        
        text = Right(text, textLength - 1)
        textLength = Len(text)
    Loop
    
    TrimLeft = text
End Function


' Trims the given character from the given text starting from the right.
' Arguments:
' - text The text to trim
' - c The character to trim
Public Function TrimRight(ByVal text As String, c As String) As String
    Dim textLength As Long
    textLength = Len(text)
    
    Dim lastCharacter As String
        
    Do While textLength > 0
        lastCharacter = Right(text, 1)
        
        If lastCharacter <> c Then
            Exit Do
        End If
        
        text = Left(text, textLength - 1)
        textLength = Len(text)
    Loop
    
    TrimRight = text
End Function


' Gets the substring from the given text.
' Arguments:
' - text The text to get the substring from
' - startIndex The index of the first character of the substring
' - [length] The amount of characters to take from the original string
Public Function Substring(ByVal text As String, ByVal startIndex As Integer, Optional ByVal length As Variant) As String
    If startIndex > Len(text) Then
        startIndex = Len(text)
    End If

    If IsMissing(length) Then
        length = Len(text) - startIndex
    End If
    
    If length > Len(text) Then
        length = Len(text) - startIndex
    End If

    Substring = Left(Right(text, Len(text) - startIndex), length)
End Function


' Checks whether the given text starts with the given sequence.
' Arguments:
' - text The text to check for the sequence
' - startText The text to be located at the start
Public Function StartsWith(ByVal text As String, ByVal startText As String) As Boolean
    StartsWith = InStr(text, startText) = 1
End Function


' Checks whether the given text ends with the given sequence.
' Arguments:
' - text The text to check for the sequence
' - endText The text to be located at the end
Public Function EndsWith(ByVal text As String, ByVal endText As String) As Boolean
    EndsWith = Right(text, Len(endText)) = endText
End Function


' Gets the character at the given index from the given string.
' Arguments:
' - text The text to get the character from
' - index The index of the character to get
Public Function CharAt(ByVal text As String, ByVal index As Integer) As String
    CharAt = Mid(text, index, 1)
End Function


' Repeats the given string the given amount of times.
' Arguments:
' - text The text to repeat
' - count The amount of times to repeat the given string
Public Function Repeat(ByVal text As String, ByVal count As Long) As String
    Repeat = ""
    
    Dim i As Long
    For i = 1 To count
        Repeat = Repeat & text
    Next
End Function


' Converts a regular string "foo" to a L"foo" string.
' Arguments:
' - text The string to convert
Public Function StringToWideString(ByVal text As String) As String
    StringToWideString = StrConv(text, vbUnicode)
End Function
