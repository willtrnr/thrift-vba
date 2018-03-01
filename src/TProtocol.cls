VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Protocol")
Option Explicit

' Transport

Public Property Get Transport() As TTransport
End Property

' Writing

Public Sub WriteMessageBegin(ByVal Message As TMessage)
End Sub

Public Sub WriteMessageEnd()
End Sub

Public Sub WriteStructBegin(ByVal Struct As TStruct)
End Sub

Public Sub WriteStructEnd()
End Sub

Public Sub WriteFieldBegin(ByVal Field As TField)
End Sub

Public Sub WriteFieldEnd()
End Sub

Public Sub WriteFieldStop()
End Sub

Public Sub WriteMapBegin(ByVal Map As TMap)
End Sub

Public Sub WriteMapEnd()
End Sub

Public Sub WriteListBegin(ByVal List As TList)
End Sub

Public Sub WriteListEnd()
End Sub

Public Sub WriteSetBegin(ByVal Set_ As TSet)
End Sub

Public Sub WriteSetEnd()
End Sub

Public Sub WriteBool(ByVal B As Boolean)
End Sub

Public Sub WriteByte(ByVal B As Byte)
End Sub

Public Sub WriteI16(ByVal I16 As Integer)
End Sub

Public Sub WriteI32(ByVal I32 As Long)
End Sub

Public Sub WriteDouble(ByVal Dub As Double)
End Sub

Public Sub WriteString(ByVal Str As String)
End Sub

Public Sub WriteBinary(ByRef Buf() As Byte)
End Sub

' Reading

Public Function ReadMessageBegin() As TMessage
End Function

Public Sub ReadMessageEnd()
End Sub

Public Function ReadStructBegin() As TStruct
End Function

Public Sub ReadStructEnd()
End Sub

Public Function ReadFieldBegin() As TField
End Function

Public Sub ReadFieldEnd()
End Sub

Public Function ReadMapBegin() As TMap
End Function

Public Sub ReadMapEnd()
End Sub

Public Function ReadListBegin() As TList
End Function

Public Sub ReadListEnd()
End Sub

Public Function ReadSetBegin() As TSet
End Function

Public Sub ReadSetEnd()
End Sub

Public Function ReadBool() As Boolean
End Function

Public Function ReadByte() As Byte
End Function

Public Function ReadI16() As Integer
End Function

Public Function ReadI32() As Long
End Function

Public Function ReadDouble() As Double
End Function

Public Function ReadString() As String
End Function

Public Function ReadBinary() As Byte()
End Function

' Reset

Public Sub Reset()
End Sub
