VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IJson

Private m_properties


Private Sub Class_Initialize()
    Set m_properties = CreateObject("Scripting.Dictionary")
End Sub


Public Property Get PropertyNames() As Collection
    Set PropertyNames = New Collection
    
    For Each property In m_properties
        PropertyNames.Add property
    Next property
End Property


Public Sub SetProperty(ByVal property As String, value As IJson)
    RemoveProperty property
    m_properties.Add property, value
End Sub


Public Sub RemoveProperty(ByVal property As String)
    If m_properties.Exists(property) Then
        m_properties.Remove property
    End If
End Sub


Public Function GetProperty(ByVal property As String) As IJson
    Set GetProperty = m_properties(property)
End Function


Private Function IJson_ToJson() As String
    Dim i As Long
    Dim propertyCount As Long
    propertyCount = m_properties.count
    
    Dim json As String
    
    For Each property In m_properties
        Dim obj As IJson
        Set obj = m_properties(property)
        
        Dim t
        t = TypeName(obj)
        
        json = json & """" & property & """: " & obj.ToJson
        
        If i < propertyCount - 1 Then
            json = json & ", "
        End If
        
        i = i + 1
    Next property
    
    IJson_ToJson = "{" & json & "}"
End Function
