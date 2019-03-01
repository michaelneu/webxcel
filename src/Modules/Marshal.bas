Attribute VB_Name = "Marshal"
Private Const integer0xFF As Integer = 256
Private Const long0xFF As Long = 256


Public Function Int8ToBytes(ByVal value As Byte) As String
    Int8ToBytes = Chr(value)
End Function


Public Function BytesToInt8(ByVal bytes As String) As Byte
    BytesToInt8 = Asc(CharAt(bytes, 1))
End Function


Public Function Int16ToBytes(ByVal value As Integer) As String
    Dim bytes As String * 2
    Dim rest As Integer

    rest = value Mod integer0xFF
    Mid(bytes, 2) = Chr(rest)

    value = (value - rest) / integer0xFF
    rest = value Mod integer0xFF
    Mid(bytes, 1) = Chr(rest)

    Int16ToBytes = bytes
End Function


Public Function BytesToInt16(ByVal bytes As String) As Integer
    BytesToInt16 = Asc(CharAt(bytes, 2)) * integer0xFF + Asc(CharAt(bytes, 1))
End Function


Public Function Int32ToBytes(ByVal value As Long) As String
    Dim bytes As String * 4
    Dim rest As Long

    rest = value Mod long0xFF
    Mid(bytes, 4) = Chr(rest)

    value = (value - rest) / long0xFF
    rest = value Mod long0xFF
    Mid(bytes, 3) = Chr(rest)

    value = (value - rest) / long0xFF
    rest = value Mod long0xFF
    Mid(bytes, 2) = Chr(rest)

    value = (value - rest) / long0xFF
    rest = value Mod long0xFF
    Mid(bytes, 1) = Chr(rest)

    Int32ToBytes = bytes
End Function


Public Function BytesToInt32(ByVal bytes As String) As Long
    BytesToInt32 = Asc(CharAt(bytes, 1)) * long0xFF * long0xFF * long0xFF + Asc(CharAt(bytes, 2)) * long0xFF * long0xFF + Asc(CharAt(bytes, 3)) * long0xFF + Asc(CharAt(bytes, 4))
End Function
