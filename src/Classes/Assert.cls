VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Indicates whether the assertion was successful
Public AssertSuccessful As Boolean
' The message of this assertion
Public AssertMessage As String


' Asserts that both values are equal. If not, this will return a failed assertion.
' Arguments:
' - expected The value the caller expects to have
' - actual The actual value
' - message A message to show in case the assertion failed
Public Function AreEqual(ByRef expected, ByRef actual, ByRef message As String) As Assert
        Set AreEqual = IsTrue(Not expected <> actual, message)
        
        If Not AreEqual.AssertSuccessful Then
            AreEqual.AssertMessage = AreEqual.AssertMessage & vbCrLf & "    Expected: " & expected & vbCrLf & "    Actual:   " & actual
        End If
End Function


' Asserts that the given value is true.
' Arguments:
' - truthy A value that can be seen as truthy. Very likely should be boolean, but maybe you like to live dangerously
' - message A message to show in case the assertion failed
Public Function IsTrue(ByRef truthy, ByRef message As String) As Assert
    Set IsTrue = New Assert
    IsTrue.AssertSuccessful = True
    IsTrue.AssertMessage = message

    If Not truthy Then
        IsTrue.AssertSuccessful = False
        IsTrue.AssertMessage = "  Assert failed: " & message
    End If
End Function


