VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGITypeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IFastCGIRecord

Public ProtocolVersion As Byte
Public MessageType As Byte


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
    Dim bytes As String
    bytes = client.ReceiveBytes(2)

    ProtocolVersion = Marshal.BytesToInt8(bytes)
    bytes = StringExtensions.Substring(bytes, 1)

    MessageType = Marshal.BytesToInt8(bytes)
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
    Dim bytes As String
    bytes = ""

    bytes = bytes & Marshal.Int8ToBytes(ProtocolVersion)
    bytes = bytes & Marshal.Int8ToBytes(MessageType)

    client.SendString bytes
End Sub
