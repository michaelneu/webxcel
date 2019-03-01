VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastCGIParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IFastCGIRecord

Public key As String
Public value As String


Private Sub IFastCGIRecord_FromBytes(ByVal bytes As String)
End Sub


Private Function IFastCGIRecord_ToBytes() As String
    Dim bytes As String
    bytes = ""
    
    bytes = bytes & Marshal.Int8ToBytes(Len(key))
    bytes = bytes & Marshal.Int8ToBytes(Len(value))
    bytes = bytes & key
    bytes = bytes & value
    
    IFastCGIRecord_ToBytes = bytes
End Function
