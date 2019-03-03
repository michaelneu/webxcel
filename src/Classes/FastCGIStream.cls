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
    m_header.ProtocolVersion = 1
    m_header.RequestId = 1
    m_header.PaddingLength = 0
End Sub


Public Property Get StreamType() As Integer
    StreamType = m_header.MessageType
End Property


Public Property Let StreamType(ByVal value As Integer)
    If value = FastCGI.FASTCGI_TYPE_STDIN Or value <> FastCGI.FASTCGI_TYPE_STDOUT Then
        m_header.MessageType = value
    End If
End Property


Private Sub IFastCGIRecord_FromBytes(ByVal bytes As String)
End Sub


Private Function IFastCGIRecord_ToBytes() As String
    m_header.ContentLength = Len(Content)

    Dim record As IFastCGIRecord
    Set record = m_header

    IFastCGIRecord_ToBytes = record.ToBytes() & Content
End Function
