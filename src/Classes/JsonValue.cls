VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IJson

Public value As Variant

Public Property Get ContainsNull() As Boolean
    ContainsNull = IsNull(value)
End Property

Private Function IJson_ToJson() As String
    If TypeName(value) = "String" Then
        IJson_ToJson = """" & value & """"
    ElseIf ContainsNull Then
        IJson_ToJson = "null"
    Else
        IJson_ToJson = value
    End If
End Function
