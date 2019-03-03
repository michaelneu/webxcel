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

Private m_header As FastCGIHeader
Public AppStatus As Long
Public ProtocolStatus As Byte
Public Reserved As String


Private Sub Class_Initialize()
    Set m_header = New FastCGIHeader
    m_header.Info.ProtocolVersion = 1
    m_header.Info.MessageType = FastCGI.FASTCGI_TYPE_END_REQUEST
    m_header.RequestId = 1
    m_header.ContentLength = 8
    m_header.PaddingLength = 0

    Reserved = StringExtensions.Repeat(Chr(0), 5)
End Sub


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
    Dim record As IFastCGIRecord
    Set record = m_header
    record.ReadFromTcpClient client

    Dim bytes As String
    bytes = client.ReceiveBytes(m_header.ContentLength)

    AppStatus = Marshal.BytesToInt32(StringExtensions.Substring(bytes, 0, 4))
    bytes = StringExtensions.Substring(bytes, 4)

    ProtocolStatus = Marshal.BytesToInt8(StringExtensions.Substring(bytes, 0, 1))
    bytes = StringExtensions.Substring(bytes, 1)

    Reserved = bytes

    client.ReceiveBytes m_header.PaddingLength
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
    Dim record As IFastCGIRecord
    Set record = m_header
    record.WriteToTcpClient client

    Dim bytes As String
    bytes = ""

    bytes = bytes & Marshal.Int32ToBytes(AppStatus)
    bytes = bytes & Marshal.Int8ToBytes(ProtocolStatus)
    bytes = bytes & Reserved

    client.SendString bytes
End Sub
