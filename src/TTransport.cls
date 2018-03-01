VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Transport")
Option Explicit

Public Function IsOpen() As Boolean
End Function

Public Function Peek() As Boolean
End Function

Public Sub TOpen()
End Sub

Public Sub TClose()
End Sub

Public Function Read(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
End Function

Public Function ReadAll(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
End Function

Public Sub TWrite(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long)
End Sub

Public Sub Flush()
End Sub
