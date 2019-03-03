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
Public Flags As Byte
Public Reserved As String


Private Sub Class_Initialize()
    Set m_header = New FastCGIHeader
    m_header.ProtocolVersion = 1
    m_header.RequestId = m_requestId
    m_header.MessageType = FastCGI.FASTCGI_TYPE_BEGIN_REQUEST
    m_header.ContentLength = 8
    m_header.PaddingLength = 0

    Role = 1
    Reserved = StringExtensions.Repeat(Chr(0), 5)
End Sub


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
    Dim header As IFastCGIRecord
    Set header = m_header

    Dim bytes As String
    bytes = header.ToBytes()

    bytes = bytes & Marshal.Int16ToBytes(Role)
    bytes = bytes & Marshal.Int8ToBytes(Flags)
    bytes = bytes & Reserved

    client.SendString bytes
End Sub
