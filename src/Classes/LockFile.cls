VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LockFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_path As String


Public Sub CreateLock(Optional ByVal name As String = "")
    If name = "" Then
        name = ActiveWorkbook.FullName & ".lock"
    End If
    
    m_path = name
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim file
    Set file = fso.CreateTextFile(m_path)
    
    file.Close
End Sub


Public Property Get Exists()
    Exists = (Not Dir(m_path) = vbNullString)
End Property
