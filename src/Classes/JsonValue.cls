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

Public Value

Private Function IJson_ToJson() As String
    If TypeName(Value) = "String" Then
        IJson_ToJson = """" & Value & """"
    Else
        IJson_ToJson = Value
    End If
End Function
