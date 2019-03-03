VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IFastCGIRecord

Private m_header As FastCGIHeader
Private m_params As Collection


Private Sub Class_Initialize()
    Set m_header = New FastCGIHeader
    m_header.ProtocolVersion = 1
    m_header.RequestId = m_requestId
    m_header.MessageType = FastCGI.FASTCGI_TYPE_PARAMS
    m_header.PaddingLength = 0

    Set m_params = New Collection
End Sub


Public Sub Add(key As String, value As String)
    Dim param As FastCGIParam
    Set param = New FastCGIParam
    
    param.key = key
    param.value = value
    
    m_params.Add param
End Sub


Private Sub IFastCGIRecord_FromBytes(ByVal bytes As String)
End Sub


Private Function IFastCGIRecord_ToBytes() As String
    Dim record As IFastCGIRecord
    Dim bytes As String
    bytes = ""

    For Each param In m_params
        Set record = param
        bytes = bytes & record.ToBytes()
    Next

    m_header.ContentLength = Len(bytes)
    Set record = m_header
    IFastCGIRecord_ToBytes = record.ToBytes() & bytes
End Function
