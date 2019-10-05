Attribute VB_Name = "TestMarshal"
Public Function TestMarshalInt8() As Assert
    Dim value As Byte
    value = 10
    
    Set TestMarshalInt8 = Assert.AreEqual(value, Marshal.BytesToInt8(Marshal.Int8ToBytes(value)), "marshals int8")
End Function

Public Function TestMarshalInt16() As Assert
    Dim value As Long
    value = 10
    
    Set TestMarshalInt16 = Assert.AreEqual(value, Marshal.BytesToInt16(Marshal.Int16ToBytes(value)), "marshals int16")
End Function

Public Function TestMarshalInt32() As Assert
    Dim value As Long
    value = 10
    
    Set TestMarshalInt32 = Assert.AreEqual(value, Marshal.BytesToInt32(Marshal.Int32ToBytes(value)), "marshals int32")
End Function
