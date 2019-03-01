VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIEndRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IFastCGIRecord

Public AppStatus As Long
Public ProtocolStatus As Byte
Public Reserved As String


Private Sub IFastCGIRecord_FromBytes(ByVal bytes As String)
End Sub


Private Function IFastCGIRecord_ToBytes() As String
End Function
