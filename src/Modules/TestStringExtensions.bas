Attribute VB_Name = "TestStringExtensions"
Public Function TestTrimLeft() As Assert
    Set TestTrimLeft = Assert.AreEqual("bar", StringExtensions.TrimLeft("oobar", "o"), "left-trims strings")
End Function


Public Function TestTrimRight() As Assert
    Set TestTrimRight = Assert.AreEqual("f", StringExtensions.TrimRight("foo", "o"), "right-trims strings")
End Function


Public Function TestStartsWith() As Assert
    Set TestStartsWith = Assert.IsTrue(StringExtensions.StartsWith("foobar", "foo"), "detects string starts")
End Function


Public Function TestEndsWith() As Assert
    Set TestEndsWith = Assert.IsTrue(StringExtensions.EndsWith("foobar", "bar"), "detects string ends")
End Function


Public Function TestCharAt() As Assert
    Set TestCharAt = Assert.AreEqual("a", StringExtensions.CharAt("foobar", 5), "gets chars from strings")
End Function


Public Function TestSubstring() As Assert
    Set TestSubstring = Assert.AreEqual("oo", StringExtensions.Substring("foobar", 1, 2), "gets parts from strings")
End Function


Public Function TestRepeat() As Assert
    Set TestRepeat = Assert.AreEqual("aaa", StringExtensions.Repeat("a", 3), "repeats strings")
End Function

