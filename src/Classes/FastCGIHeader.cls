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

Private m_typeInfo As FastCGITypeInfo
Public RequestId As Integer
Public ContentLength As Long
Public PaddingLength As Byte
Public Reserved As String


Private Sub Class_Initialize()
    Set m_typeInfo = New FastCGITypeInfo
    Reserved = StringExtensions.Repeat(Chr(0), 1)
End Sub


Public Property Get Info() As FastCGITypeInfo
    Set Info = m_typeInfo
End Property


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
    Dim bytes As String
    bytes = client.ReceiveBytes(6)

    RequestId = Marshal.BytesToInt16(bytes)
    bytes = StringExtensions.Substring(bytes, 2)

    ContentLength = Marshal.BytesToInt16(bytes)
    bytes = StringExtensions.Substring(bytes, 2)

    PaddingLength = Marshal.BytesToInt8(bytes)
    bytes = StringExtensions.Substring(bytes, 1)

    Reserved = bytes
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
    Dim bytes As String
    bytes = ""

    bytes = bytes & Marshal.Int8ToBytes(m_typeInfo.ProtocolVersion)
    bytes = bytes & Marshal.Int8ToBytes(m_typeInfo.MessageType)
    bytes = bytes & Marshal.Int16ToBytes(RequestId)
    bytes = bytes & Marshal.Int16ToBytes(ContentLength)
    bytes = bytes & Marshal.Int8ToBytes(PaddingLength)
    bytes = bytes & Reserved

    client.SendString bytes
End Sub
