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
    m_header.Info.ProtocolVersion = 1
    m_header.Info.MessageType = FastCGI.FASTCGI_TYPE_PARAMS
    m_header.RequestId = 1
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


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
    Dim record As IFastCGIRecord
    Set record = m_header
    record.ReadFromTcpClient client

    Set m_params = New Collection

    Dim bytes As String
    bytes = client.ReceiveBytes(m_header.ContentLength)
    
    Dim keyLength As Integer
    Dim valueLength As Integer
    Dim param As FastCGIParam

    Do While Len(bytes) > 0
        keyLength = Marshal.BytesToInt8(StringExtensions.Substring(bytes, 0, 1))
        bytes = StringExtensions.Substring(bytes, 1)
        valueLength = Marshal.BytesToInt8(StringExtensions.Substring(bytes, 0, 1))
        bytes = StringExtensions.Substring(bytes, 1)
        
        Set param = New FastCGIParam

        param.key = StringExtensions.Substring(bytes, 0, keyLength)
        bytes = StringExtensions.Substring(bytes, keyLength)

        param.value = StringExtensions.Substring(bytes, 0, valueLength)
        bytes = StringExtensions.Substring(bytes, valueLength)

        m_params.Add param
    Loop

    client.ReceiveBytes m_header.PaddingLength
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
    Dim record As IFastCGIRecord
    Dim bytes As String
    bytes = ""

    For Each param In m_params
        bytes = bytes & Marshal.Int8ToBytes(Len(param.key))
        bytes = bytes & Marshal.Int8ToBytes(Len(param.value))
        bytes = bytes & param.key
        bytes = bytes & param.value
    Next

    m_header.ContentLength = Len(bytes)
    Set record = m_header
    record.WriteToTcpClient client
    client.SendString bytes
End Sub
