VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String
Private m_value As String


Public Property Get name() As String
    name = m_name
End Property


Public Property Let name(ByVal Value As String)
    m_name = Value
End Property


Public Property Get Value() As String
    Value = m_value
End Property


Public Property Let Value(ByVal val As String)
    m_value = val
End Property
