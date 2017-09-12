VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TcpError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub RaiseErrorSetup()
    Err.Raise (-1)
End Sub

Public Sub RaiseErrorCreateSocket()
    Err.Raise (-2)
End Sub

Public Sub RaiseErrorBind(ByVal port As Long)
    Err.Raise (-3)
End Sub

Public Sub RaiseErrorListen()
    Err.Raise (-4)
End Sub

Public Sub RaiseErrorAccept()
    Err.Raise (-5)
End Sub
