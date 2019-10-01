VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIBeginRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IFastCGIRecord

Private m_header As FastCGIHeader
Public Role As Integer
Public flags As Byte
Public Reserved As String


Private Sub Class_Initialize()
    Set m_header = New FastCGIHeader
    m_header.Info.ProtocolVersion = 1
    m_header.Info.MessageType = FastCGI.FASTCGI_TYPE_BEGIN_REQUEST
    m_header.RequestId = 1
    m_header.ContentLength = 8
    m_header.PaddingLength = 0

    Role = 1
    Reserved = StringExtensions.Repeat(Chr(0), 5)
End Sub


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
    Dim record As IFastCGIRecord
    Set record = m_header
    record.ReadFromTcpClient client

    Dim bytes As String
    bytes = client.ReceiveBytes(m_header.ContentLength)

    Role = Marshal.BytesToInt16(StringExtensions.Substring(bytes, 0, 2))
    bytes = StringExtensions.Substring(bytes, 2)

    flags = Marshal.BytesToInt8(StringExtensions.Substring(bytes, 0, 1))
    bytes = StringExtensions.Substring(bytes, 1)

    Reserved = bytes

    client.ReceiveBytes m_header.PaddingLength
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
    Dim header As IFastCGIRecord
    Set header = m_header
    header.WriteToTcpClient client

    Dim bytes As String
    bytes = ""

    bytes = bytes & Marshal.Int16ToBytes(Role)
    bytes = bytes & Marshal.Int8ToBytes(flags)
    bytes = bytes & Reserved

    client.SendString bytes
End Sub
