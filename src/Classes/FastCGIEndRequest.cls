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
End Sub


Private Sub IFastCGIRecord_ReadFromTcpClient(client As TcpClient)
End Sub


Private Sub IFastCGIRecord_WriteToTcpClient(client As TcpClient)
End Sub
