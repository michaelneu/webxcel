Attribute VB_Name = "TestUtils"
Public Type TestFileHandle
    TestFilename As String
    TestFilestream As Variant
End Type

' Creates a temporary test file.
Public Function CreateTestFile() As TestFileHandle
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CreateTestFile.TestFilename = fso.BuildPath(ActiveDocument.Path, fso.GetTempName())
    Set CreateTestFile.TestFilestream = fso.CreateTextFile(CreateTestFile.TestFilename, True)
End Function


' Deletes the given test file.
' Arguments:
' - fileHandle The file handle to delete
Public Sub DeleteTestFile(ByRef fileHandle As TestFileHandle)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFile fileHandle.TestFilename
End Sub
