VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TBinaryProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Protocol")
Option Explicit

Implements TProtocol

Private Const CP_UTF8 As Long = 65001

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
  ByVal CodePage As Long, _
  ByVal dwFlags As Long, _
  ByVal lpWideCharStr As LongPtr, _
  ByVal cchWideChar As Long, _
  ByVal lpMultiByteStr As LongPtr, _
  ByVal cbMultiByte As Long, _
  ByVal lpDefaultChar As LongPtr, _
  ByVal lpUsedDefaultChar As LongPtr _
) As Long

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
  ByVal CodePage As Long, _
  ByVal dwFlags As Long, _
  ByVal lpMultiByteStr As LongPtr, _
  ByVal cbMultiByte As Long, _
  ByVal lpWideCharStr As LongPtr, _
  ByVal cchWideChar As Long _
) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private Const NO_LENGTH_LIMIT As Long = -1

Private Const VERSION_MASK As Long = &HFFFF0000
Private Const VERSION_1 As Long = &H80010000

Private m_init As Boolean

Private m_trans As TTransport
Private m_strictRead As Boolean
Private m_strictWrite As Boolean
Private m_stringLengthLimit As Long
Private m_containerLengthLimit As Long

Private Sub Class_Initialize()
  m_init = False
  Set m_trans = Nothing
  m_strictRead = False
  m_strictWrite = False
End Sub

Public Sub Init(ByVal Trans As TTransport, Optional ByVal StrictRead As Boolean = False, Optional ByVal StrictWrite As Boolean = True, Optional ByVal StringLengthLimit As Long = NO_LENGTH_LIMIT, Optional ByVal ContainerLengthLimit As Long = NO_LENGTH_LIMIT)
  Debug.Assert Not m_init
  m_init = True
  Set m_trans = Trans
  m_strictRead = StrictRead
  m_strictWrite = StrictWrite
  m_stringLengthLimit = StringLengthLimit
  m_containerLengthLimit = ContainerLengthLimit
End Sub

' Transport

Public Property Get Transport() As TTransport
  Debug.Assert m_init
  Set Transport = m_trans
End Property

' Writing

Public Sub WriteMessageBegin(ByVal Message As TMessage)
  Debug.Assert m_init
  If m_strictWrite Then
    WriteI32 (VERSION_1 Or CLng(Message.TType))
    WriteString Message.Name
    WriteI32 Message.SeqId
  Else
    WriteString Message.Name
    WriteByte Message.TType
    WriteI32 Message.SeqId
  End If
End Sub

Public Sub WriteMessageEnd()
  Debug.Assert m_init
End Sub

Public Sub WriteStructBegin(ByVal Struct As TStruct)
  Debug.Assert m_init
End Sub

Public Sub WriteStructEnd()
  Debug.Assert m_init
End Sub

Public Sub WriteFieldBegin(ByVal Field As TField)
  Debug.Assert m_init
  WriteByte Field.TType
  WriteI16 Field.Id
End Sub

Public Sub WriteFieldEnd()
  Debug.Assert m_init
End Sub

Public Sub WriteFieldStop()
  Debug.Assert m_init
  WriteByte TType_Stop
End Sub

Public Sub WriteMapBegin(ByVal Map As TMap)
  Debug.Assert m_init
  WriteByte Map.KeyType
  WriteByte Map.ValueType
  WriteI32 Map.Size
End Sub

Public Sub WriteMapEnd()
  Debug.Assert m_init
End Sub

Public Sub WriteListBegin(ByVal List As TList)
  Debug.Assert m_init
  WriteByte List.ElemType
  WriteI32 List.Size
End Sub

Public Sub WriteListEnd()
  Debug.Assert m_init
End Sub

Public Sub WriteSetBegin(ByVal Set_ As TSet)
  Debug.Assert m_init
  WriteByte Set_.ElemType
  WriteI32 Set_.Size
End Sub

Public Sub WriteSetEnd()
  Debug.Assert m_init
End Sub

Public Sub WriteBool(ByVal B As Boolean)
  Debug.Assert m_init
  If B Then
    WriteByte 0
  Else
    WriteByte 1
  End If
End Sub

Public Sub WriteByte(ByVal B As Byte)
  Debug.Assert m_init
  Dim Buffer(0) As Byte: Buffer(0) = B
  m_trans.TWrite Buffer, 0, 1
End Sub

Public Sub WriteI16(ByVal I16 As Integer)
  Debug.Assert m_init
  Dim Buffer(1) As Byte
  Buffer(0) = CByte((I16 And &HFF00&) \ &H100&)
  Buffer(1) = CByte(I16 And &HFF&)
  m_trans.TWrite Buffer, 0, 2
End Sub

Public Sub WriteI32(ByVal I32 As Long)
  Debug.Assert m_init
  Dim Buffer(3) As Byte
  Buffer(0) = CByte((I32 And &H7F000000) \ &H1000000)
  Buffer(1) = CByte((I32 And &HFF0000) \ &H10000)
  Buffer(2) = CByte((I32 And &HFF00&) \ &H100&)
  Buffer(3) = CByte(I32 And &HFF&)
  If I32 < 0 Then
    Buffer(0) = Buffer(0) Or &H80
  End If
  m_trans.TWrite Buffer, 0, 4
End Sub

Public Sub WriteI64(ByVal I64 As TLongLong)
  Debug.Assert m_init
  m_trans.TWrite I64.Bytes, 0, 8
End Sub

Public Sub WriteDouble(ByVal Dub As Double)
  Debug.Assert m_init
  Dim Buffer(7) As Byte
  CopyMemory VarPtr(Buffer(0)), VarPtr(Dub), 8
  m_trans.TWrite Buffer, 0, 8
End Sub

Public Sub WriteString(ByVal Str As String)
  Debug.Assert m_init
  If Str = vbNullString Then
      WriteI32 0
  Else
    Dim Buffer() As Byte
    
    Dim Size As Long: Size = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(Str), -1&, 0&, 0&, 0&, 0&)
    ReDim Buffer(Size - 1)
    Size = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(Str), -1&, VarPtr(Buffer(0)), Size, 0&, 0&)
    
    WriteI32 Size - 1
    m_trans.TWrite Buffer, 0, Size - 1
  End If
End Sub

Public Sub WriteBinary(ByRef Buffer() As Byte)
  Debug.Assert m_init
  WriteI32 UBound(Buffer) + 1
  m_trans.TWrite Buffer, 0, UBound(Buffer) + 1
End Sub

' Reading

Public Function ReadMessageBegin() As TMessage
  Debug.Assert m_init
  
  Dim Size As Long
  Size = ReadI32
  
  Dim Name As String
  Dim TType As Byte
  
  If Size < 0 Then
    If (Size And VERSION_MASK) <> VERSION_1 Then
      Err.Raise 5, Description:="Bad version in ReadMessageBegin"
    End If
    Name = ReadString
    TType = CByte(Size And &HFF&)
  ElseIf m_strictRead Then
    Err.Raise 5, Description:="Missing version in ReadMessageBegin, old client?"
  Else
    CheckStringLength Size
    Name = ReadStringBody(Size)
    TType = ReadByte
  End If
  
  Dim SeqId As Long: SeqId = ReadI32
  Set ReadMessageBegin = NewTMessage(Name, TType, SeqId)
End Function

Public Sub ReadMessageEnd()
  Debug.Assert m_init
End Sub

Public Function ReadStructBegin() As TStruct
  Debug.Assert m_init
  Set ReadStructBegin = NewTStruct(vbNullString)
End Function

Public Sub ReadStructEnd()
  Debug.Assert m_init
End Sub

Public Function ReadFieldBegin() As TField
  Debug.Assert m_init
  
  Dim TType As Byte: TType = ReadByte

  Dim Id As Integer
  If TType = TType_Stop Then
    Id = 0
  Else
    Id = ReadI16
  End If
  
  Set ReadFieldBegin = NewTField(vbNullString, TType, Id)
End Function

Public Sub ReadFieldEnd()
  Debug.Assert m_init
End Sub

Public Function ReadMapBegin() As TMap
  Debug.Assert m_init
  
  Dim KType As Byte: KType = ReadByte
  Dim VType As Byte: VType = ReadByte
  Dim Size As Long: Size = ReadI32
  CheckContainerLength Size
  
  Set ReadMapBegin = NewTMap(KType, VType, Size)
End Function

Public Sub ReadMapEnd()
  Debug.Assert m_init
End Sub

Public Function ReadListBegin() As TList
  Debug.Assert m_init
  
  Dim TType As Byte: TType = ReadByte
  Dim Size As Long: Size = ReadI32
  CheckContainerLength Size
  
  Set ReadListBegin = NewTList(TType, Size)
End Function

Public Sub ReadListEnd()
  Debug.Assert m_init
End Sub

Public Function ReadSetBegin() As TSet
  Debug.Assert m_init
  
  Dim TType As Byte: TType = ReadByte
  Dim Size As Long: Size = ReadI32
  CheckContainerLength Size
  
  Set ReadSetBegin = NewTSet(TType, Size)
End Function

Public Sub ReadSetEnd()
  Debug.Assert m_init
End Sub

Public Function ReadBool() As Boolean
  Debug.Assert m_init
  ReadBool = ReadByte = 1
End Function

Public Function ReadByte() As Byte
  Debug.Assert m_init
  
  Dim Buffer(0) As Byte
  m_trans.ReadAll Buffer, 0, 1
  
  ReadByte = Buffer(0)
End Function

Public Function ReadI16() As Integer
  Debug.Assert m_init
  
  Dim Buffer(1) As Byte
  m_trans.ReadAll Buffer, 0, 2
  
  ReadI16 = CLng(Buffer(0) And &H7F) * &H100 + CLng(Buffer(1))
  If (Buffer(0) And &H80) = &H80 Then
    ReadI16 = ReadI16 Or &H8000
  End If
End Function

Public Function ReadI32() As Long
  Debug.Assert m_init
  
  Dim Buffer(3) As Byte
  m_trans.ReadAll Buffer, 0, 4
  
  ReadI32 = CLng(Buffer(0) And &H7F) * &H1000000 + CLng(Buffer(1)) * &H10000 + CLng(Buffer(2)) * &H100 + CLng(Buffer(3))
  If (Buffer(0) And &H80) = &H80 Then
    ReadI32 = ReadI32 Or &H80000000
  End If
End Function

Public Function ReadI64() As TLongLong
  Debug.Assert m_init

  Dim Buffer(7) As Byte
  m_trans.ReadAll Buffer, 0, 8

  Set ReadI64 = NewTLongLong(Buffer)
End Function

Public Function ReadDouble() As Double
  Debug.Assert m_init
  
  Dim Buffer(7) As Byte
  m_trans.ReadAll Buffer, 0, 8
  
  CopyMemory VarPtr(ReadDouble), VarPtr(Buffer(0)), 8
End Function

Private Function ReadStringBody(ByVal Size As Long) As String
  Debug.Assert m_init
  If Size = 0 Then
    ReadStringBody = vbNullString
  Else
    Dim Buffer() As Byte: ReDim Buffer(Size - 1)
    m_trans.ReadAll Buffer, 0, Size
    
    Dim StrSize As Long: StrSize = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(Buffer(0)), Size, 0&, 0&)
    Dim StrBuffer() As Byte: ReDim StrBuffer(StrSize * 2 - 1)
    MultiByteToWideChar CP_UTF8, 0&, VarPtr(Buffer(0)), Size, VarPtr(StrBuffer(0)), StrSize

    ReadStringBody = StrBuffer
  End If
End Function

Public Function ReadString() As String
  Debug.Assert m_init
  
  Dim Size As Long: Size = ReadI32
  CheckStringLength Size

  ReadString = ReadStringBody(Size)
End Function

Public Function ReadBinary() As Byte()
  Debug.Assert m_init
  
  Dim Size As Long: Size = ReadI32
  CheckStringLength Size
  
  Dim Buffer() As Byte: ReDim Buffer(Size - 1)
  m_trans.ReadAll Buffer, 0, Size

  ReadBinary = Buffer
End Function

Private Sub CheckStringLength(ByVal Size As Long)
  If Size < 0 Then
    Err.Raise 5, Description:="Negative length"
  ElseIf m_stringLengthLimit <> NO_LENGTH_LIMIT And Size > m_stringLengthLimit Then
    Err.Raise 5, Description:="Length exceeded max allowed"
  End If
End Sub

Private Sub CheckContainerLength(ByVal Size As Long)
  If Size < 0 Then
    Err.Raise 5, Description:="Negative length"
  ElseIf m_containerLengthLimit <> NO_LENGTH_LIMIT And Size > m_containerLengthLimit Then
    Err.Raise 5, Description:="Length exceeded max allowed"
  End If
End Sub

' Reset

Public Sub Reset()
  Debug.Assert m_init
End Sub


' Interface implementation

Private Property Get TProtocol_Transport() As TTransport
  Set TProtocol_Transport = Transport
End Property

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

Private Sub TProtocol_WriteI64(ByVal I64 As TLongLong)
  WriteI64 I64
End Sub

Private Sub TProtocol_WriteDouble(ByVal Dub As Double)
  WriteDouble Dub
End Sub

Private Sub TProtocol_WriteString(ByVal Str As String)
  WriteString Str
End Sub

Private Sub TProtocol_WriteBinary(ByRef Buffer() As Byte)
  WriteBinary Buffer
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

Private Function TProtocol_ReadI64() As TLongLong
  Set TProtocol_ReadI64 = ReadI64
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


