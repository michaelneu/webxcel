VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IJson

Private m_items As Collection


Private Sub Class_Initialize()
    Set m_items = New Collection
End Sub


Public Property Get count() As Long
    count = m_items.count
End Property


Public Sub AddItem(item As IJson)
    m_items.Add item
End Sub


Public Function GetItem(index As Integer) As IJson
    Set GetItem = m_items(index + 1)
End Function


Private Function IJson_ToJson() As String
    Dim json As String
    Dim i As Integer
    
    For i = 1 To count
        Dim obj As IJson
        Set obj = m_items(i)
        
        json = json & obj.ToJson
        
        If i < count Then
            json = json & ", "
        End If
    Next i
    
    IJson_ToJson = "[" & json & "]"
End Function
