VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IFastCGIRecord

Public ProtocolVersion As Byte
Public MessageType As Byte
Public RequestId As Integer
Public ContentLength As Integer
Public PaddingLength As Byte
Public Reserved As String


Private Sub Class_Initialize()
    Dim bytes As String * 1
    Reserved = bytes
End Sub


Private Sub IFastCGIRecord_FromBytes(ByVal bytes As String)
End Sub


Private Function IFastCGIRecord_ToBytes() As String
    Dim bytes As String
    bytes = ""

    bytes = bytes & Marshal.Int8ToBytes(ProtocolVersion)
    bytes = bytes & Marshal.Int8ToBytes(MessageType)
    bytes = bytes & Marshal.Int16ToBytes(RequestId)
    bytes = bytes & Marshal.Int16ToBytes(ContentLength)
    bytes = bytes & Marshal.Int8ToBytes(PaddingLength)
    bytes = bytes & Reserved

    IFastCGIRecord_ToBytes = bytes
End Function
