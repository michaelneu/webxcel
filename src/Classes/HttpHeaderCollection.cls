VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpHeaderCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_headers


Private Sub Class_Initialize()
    Set m_headers = CreateObject("Scripting.Dictionary")
End Sub


Public Sub AddHeader(ByVal name As String, ByVal Value As String)
    If m_headers.Exists(name) Then
        m_headers.Remove name
    End If
    
    m_headers.Add name, Value
End Sub


Public Function GetEnumerator() As Collection
    Dim headerCollection As Collection
    Set headerCollection = New Collection
    
    For Each Key In m_headers.Keys
        Dim header As HttpHeader
        Set header = New HttpHeader
        
        header.name = Key
        header.Value = m_headers(Key)
        
        headerCollection.Add header
    Next Key
    
    Set GetEnumerator = headerCollection
End Function

