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


Public Sub AddHeader(ByVal name As String, ByVal value As String)
    If m_headers.Exists(name) Then
        m_headers.Remove name
    End If
    
    m_headers.Add name, value
End Sub


Public Function GetEnumerator() As Collection
    Dim headerCollection As Collection
    Set headerCollection = New Collection
    
    For Each key In m_headers.Keys
        Dim header As HttpHeader
        Set header = New HttpHeader
        
        header.name = key
        header.value = m_headers(key)
        
        headerCollection.Add header
    Next key
    
    Set GetEnumerator = headerCollection
End Function

