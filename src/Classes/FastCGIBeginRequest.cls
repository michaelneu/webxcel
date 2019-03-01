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

Public Role As Integer
Public Flags As Byte
Public Reserved As String


Private Sub Class_Initialize()
    Dim bytes As String * 5
    Reserved = bytes
    Role = 1
End Sub



Private Sub IFastCGIRecord_FromBytes(ByVal bytes As String)
End Sub


Private Function IFastCGIRecord_ToBytes() As String
    Dim bytes As String
    bytes = ""

    bytes = bytes & Marshal.Int16ToBytes(Role)
    bytes = bytes & Marshal.Int8ToBytes(Flags)
    bytes = bytes & Reserved

    IFastCGIRecord_ToBytes = bytes
End Function
