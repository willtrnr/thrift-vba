VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TMultiplexedProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Protocol")
Option Explicit

Implements TProtocol

Private Const SEPARATOR As String = ":"

Private m_init As Boolean

Private m_proto As TProtocol
Private m_serviceName As String

Private Sub Class_Initialize()
  m_init = False
  Set m_proto = Nothing
  m_serviceName = vbNullString
End Sub

Public Sub Init(ByVal Proto As TProtocol, ByVal ServiceName As String)
  Debug.Assert Not m_init
  m_init = True
  Set m_proto = Proto
  m_serviceName = ServiceName
End Sub

' Transport

Public Function GetTransport() As TTransport
  Debug.Assert m_init
  Set GetTransport = m_proto.GetTransport
End Function

' Writing

Public Sub WriteMessageBegin(ByVal Message As TMessage)
  Debug.Assert m_init
  If Message.TType = TMessageType_Call Or Message.TType = TMessageType_OneWay Then
    Dim M As TMessage
    Set M = New TMessage
    M.Init m_serviceName & SEPARATOR & Message.Name, Message.TType, Message.SeqId
    m_proto.WriteMessageBegin M
  Else
    m_proto.WriteMessageBegin Message
  End If
End Sub

Public Sub WriteMessageEnd()
  Debug.Assert m_init
  m_proto.WriteMessageEnd
End Sub

Public Sub WriteStructBegin(ByVal Struct As TStruct)
  Debug.Assert m_init
  m_proto.WriteStructBegin Struct
End Sub

Public Sub WriteStructEnd()
  Debug.Assert m_init
  m_proto.WriteStructEnd
End Sub

Public Sub WriteFieldBegin(ByVal Field As TField)
  Debug.Assert m_init
  m_proto.WriteFieldBegin Field
End Sub

Public Sub WriteFieldEnd()
  Debug.Assert m_init
  m_proto.WriteFieldEnd
End Sub

Public Sub WriteFieldStop()
  Debug.Assert m_init
  m_proto.WriteFieldStop
End Sub

Public Sub WriteMapBegin(ByVal Map As TMap)
  Debug.Assert m_init
  m_proto.WriteMapBegin Map
End Sub

Public Sub WriteMapEnd()
  Debug.Assert m_init
  m_proto.WriteMapEnd
End Sub

Public Sub WriteListBegin(ByVal List As TList)
  Debug.Assert m_init
  m_proto.WriteListBegin List
End Sub

Public Sub WriteListEnd()
  Debug.Assert m_init
  m_proto.WriteListEnd
End Sub

Public Sub WriteSetBegin(ByVal Set_ As TSet)
  Debug.Assert m_init
  m_proto.WriteSetBegin Set_
End Sub

Public Sub WriteSetEnd()
  Debug.Assert m_init
  m_proto.WriteSetEnd
End Sub

Public Sub WriteBool(ByVal B As Boolean)
  Debug.Assert m_init
  m_proto.WriteBool B
End Sub

Public Sub WriteByte(ByVal B As Byte)
  Debug.Assert m_init
  m_proto.WriteByte B
End Sub

Public Sub WriteI16(ByVal I16 As Integer)
  Debug.Assert m_init
  m_proto.WriteI16 I16
End Sub

Public Sub WriteI32(ByVal I32 As Long)
  Debug.Assert m_init
  m_proto.WriteI32 I32
End Sub

Public Sub WriteDouble(ByVal Dub As Double)
  Debug.Assert m_init
  m_proto.WriteDouble Dub
End Sub

Public Sub WriteString(ByVal Str As String)
  Debug.Assert m_init
  m_proto.WriteString Str
End Sub

Public Sub WriteBinary(ByRef Buf() As Byte)
  Debug.Assert m_init
  m_proto.WriteBinary Buf
End Sub

' Reading

Public Function ReadMessageBegin() As TMessage
  Debug.Assert m_init
  Set ReadMessageBegin = m_proto.ReadMessageBegin
End Function

Public Sub ReadMessageEnd()
  Debug.Assert m_init
  m_proto.ReadMessageEnd
End Sub

Public Function ReadStructBegin() As TStruct
  Debug.Assert m_init
  Set ReadStructBegin = m_proto.ReadStructBegin
End Function

Public Sub ReadStructEnd()
  Debug.Assert m_init
  m_proto.ReadStructEnd
End Sub

Public Function ReadFieldBegin() As TField
  Debug.Assert m_init
  Set ReadFieldBegin = m_proto.ReadFieldBegin
End Function

Public Sub ReadFieldEnd()
  Debug.Assert m_init
  m_proto.ReadFieldEnd
End Sub

Public Function ReadMapBegin() As TMap
  Debug.Assert m_init
  Set ReadMapBegin = m_proto.ReadMapBegin
End Function

Public Sub ReadMapEnd()
  Debug.Assert m_init
  m_proto.ReadMapEnd
End Sub

Public Function ReadListBegin() As TList
  Debug.Assert m_init
  Set ReadListBegin = m_proto.ReadListBegin
End Function

Public Sub ReadListEnd()
  Debug.Assert m_init
  m_proto.ReadListEnd
End Sub

Public Function ReadSetBegin() As TSet
  Debug.Assert m_init
  Set ReadSetBegin = m_proto.ReadSetBegin
End Function

Public Sub ReadSetEnd()
  Debug.Assert m_init
  m_proto.ReadSetEnd
End Sub

Public Function ReadBool() As Boolean
  Debug.Assert m_init
  ReadBool = m_proto.ReadBool
End Function

Public Function ReadByte() As Byte
  Debug.Assert m_init
  ReadByte = m_proto.ReadByte
End Function

Public Function ReadI16() As Integer
  Debug.Assert m_init
  ReadI16 = m_proto.ReadI16
End Function

Public Function ReadI32() As Long
  Debug.Assert m_init
  ReadI32 = m_proto.ReadI32
End Function

Public Function ReadDouble() As Double
  Debug.Assert m_init
  ReadDouble = m_proto.ReadDouble
End Function

Public Function ReadString() As String
  Debug.Assert m_init
  ReadString = m_proto.ReadString
End Function

Public Function ReadBinary() As Byte()
  Debug.Assert m_init
  ReadBinary = m_proto.ReadBinary
End Function

' Reset

Public Sub Reset()
  Debug.Assert m_init
  m_proto.Reset
End Sub


' Interface implementation

Private Function TProtocol_GetTransport() As TTransport
  Set TProtocol_GetTransport = GetTransport()
End Function

Private Sub TProtocol_WriteMessageBegin(ByVal Message As TMessage)
  WriteMessageBegin Message
End Sub

Private Sub TProtocol_WriteMessageEnd()
  WriteMessageEnd
End Sub

Private Sub TProtocol_WriteStructBegin(ByVal Struct As TStruct)
  WriteStructBegin Struct
End Sub

Private Sub TProtocol_WriteStructEnd()
  WriteStructEnd
End Sub

Private Sub TProtocol_WriteFieldBegin(ByVal Field As TField)
  WriteFieldBegin Field
End Sub

Private Sub TProtocol_WriteFieldEnd()
  WriteFieldEnd
End Sub

Private Sub TProtocol_WriteFieldStop()
  WriteFieldStop
End Sub

Private Sub TProtocol_WriteMapBegin(ByVal Map As TMap)
  WriteMapBegin Map
End Sub

Private Sub TProtocol_WriteMapEnd()
  WriteMapEnd
End Sub

Private Sub TProtocol_WriteListBegin(ByVal List As TList)
  WriteListBegin List
End Sub

Private Sub TProtocol_WriteListEnd()
  WriteListEnd
End Sub

Private Sub TProtocol_WriteSetBegin(ByVal Set_ As TSet)
  WriteSetBegin Set_
End Sub

Private Sub TProtocol_WriteSetEnd()
  WriteSetEnd
End Sub

Private Sub TProtocol_WriteBool(ByVal B As Boolean)
  WriteBool B
End Sub

Private Sub TProtocol_WriteByte(ByVal B As Byte)
  WriteByte B
End Sub

Private Sub TProtocol_WriteI16(ByVal I16 As Integer)
  WriteI16 I16
End Sub

Private Sub TProtocol_WriteI32(ByVal I32 As Long)
  WriteI32 I32
End Sub

Private Sub TProtocol_WriteDouble(ByVal Dub As Double)
  WriteDouble Dub
End Sub

Private Sub TProtocol_WriteString(ByVal Str As String)
  WriteString Str
End Sub

Private Sub TProtocol_WriteBinary(ByRef Buf() As Byte)
  WriteBinary Buf
End Sub

Private Function TProtocol_ReadMessageBegin() As TMessage
  Set TProtocol_ReadMessageBegin = ReadMessageBegin
End Function

Private Sub TProtocol_ReadMessageEnd()
  ReadMessageEnd
End Sub

Private Function TProtocol_ReadStructBegin() As TStruct
  Set TProtocol_ReadStructBegin = ReadStructBegin
End Function

Private Sub TProtocol_ReadStructEnd()
  ReadStructEnd
End Sub

Private Function TProtocol_ReadFieldBegin() As TField
  Set TProtocol_ReadFieldBegin = ReadFieldBegin
End Function

Private Sub TProtocol_ReadFieldEnd()
  ReadFieldEnd
End Sub

Private Function TProtocol_ReadMapBegin() As TMap
  Set TProtocol_ReadMapBegin = ReadMapBegin
End Function

Private Sub TProtocol_ReadMapEnd()
  ReadMapEnd
End Sub

Private Function TProtocol_ReadListBegin() As TList
  Set TProtocol_ReadListBegin = ReadListBegin
End Function

Private Sub TProtocol_ReadListEnd()
  ReadListEnd
End Sub

Private Function TProtocol_ReadSetBegin() As TSet
  Set TProtocol_ReadSetBegin = ReadSetBegin
End Function

Private Sub TProtocol_ReadSetEnd()
  ReadSetEnd
End Sub

Private Function TProtocol_ReadBool() As Boolean
  TProtocol_ReadBool = ReadBool
End Function

Private Function TProtocol_ReadByte() As Byte
  TProtocol_ReadByte = ReadByte
End Function

Private Function TProtocol_ReadI16() As Integer
  TProtocol_ReadI16 = ReadI16
End Function

Private Function TProtocol_ReadI32() As Long
  TProtocol_ReadI32 = ReadI32
End Function

Private Function TProtocol_ReadDouble() As Double
  TProtocol_ReadDouble = ReadDouble
End Function

Private Function TProtocol_ReadString() As String
  TProtocol_ReadString = ReadString
End Function

Private Function TProtocol_ReadBinary() As Byte()
  TProtocol_ReadBinary = ReadBinary
End Function

Public Sub TProtocol_Reset()
  Reset
End Sub

