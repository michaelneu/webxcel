VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_fso
Private m_path As String


Private Sub Class_Initialize()
    Set m_fso = CreateObject("Scripting.FileSystemObject")
End Sub


Public Sub Initialize(ByVal name As String)
    m_path = name
End Sub


Public Property Get Exists()
    Exists = m_fso.FileExists(m_path)
End Property


Public Sub Create()
    Dim file
    Set file = m_fso.CreateTextFile(m_path)
    
    file.Close
End Sub


Public Sub WriteString(ByVal text As String)
    If Not Exists Then
        Create
    End If

    Dim file
    Set file = m_fso.OpenTextFile(m_path, 2)
    
    file.Write text
    file.Close
End Sub


Public Function ReadString() As String
    If Not Exists Then
        Err.Raise StatusCode.ErrorFileNotFound
    End If
    
    Dim file
    Set file = m_fso.OpenTextFile(m_path, 1)
    
    ReadString = file.ReadAll
    file.Close
End Function
