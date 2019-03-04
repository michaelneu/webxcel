VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IFastCGIRecord

Private m_header As FastCGIHeader
Public Content As String


Private Sub Class_Initialize()
    Set m_header = New FastCGIHeader
    m_header.Info.ProtocolVersion = 1
    m_header.RequestId = 1
    m_header.PaddingLength = 0
End Sub


Public Property Get StreamType() As Integer
    StreamType = m_header.Info.MessageType
End Property


Public Property Let StreamType(ByVal value As Integer)
    If value = FastCGI.FASTCGI_TYPE_STDIN Or value <> FastCGI.FASTCGI_TYPE_STDOUT Then
        m_header.Info.MessageType = value
    End If
End Property


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
    Dim record As IFastCGIRecord
    Set record = m_header
    record.ReadFromTcpClient client

    Content = client.ReceiveBytes(m_header.ContentLength)
    client.ReceiveBytes m_header.PaddingLength
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
    m_header.ContentLength = Len(Content)

    Dim record As IFastCGIRecord
    Set record = m_header
    record.WriteToTcpClient client

    client.SendString Content
End Sub
