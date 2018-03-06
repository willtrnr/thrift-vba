VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TApplicationException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift")
Option Explicit

Implements TSerializable

Private m_init As Boolean

Private m_message As String
Private m_type As Long

Private Sub Class_Initialize()
  m_init = False
  m_message = vbNullString
  m_type = 0
End Sub

Public Sub Init(Optional ByVal Message As String = vbNullString, Optional ByVal TType As Byte = 0)
  Debug.Assert Not m_init
  m_init = True
  m_type = TType
  m_message = Message
End Sub

Public Sub TRead(ByVal Proto As TProtocol)
  Debug.Assert Not m_init
  
  Dim TType As Byte: TType = 0
  Dim Message As String: Message = vbNullString
  
  Proto.ReadStructBegin
  
  Dim Field As TField
  Do While True
    Set Field = Proto.ReadFieldBegin
    If Field.TType = TType_Stop Then
      Exit Do
    ElseIf Field.Id = 1 And Field.TType = TType_String Then
      Message = Proto.ReadString
    ElseIf Field.Id = 2 And Field.TType = TType_I32 Then
      TType = Proto.ReadI32
    Else
      TProtocolUtil.Skip Proto, Field.TType
    End If
    Proto.ReadFieldEnd
  Loop
  
  Proto.ReadStructEnd
  
  Init Message, TType
End Sub

Public Sub TWrite(ByVal Proto As TProtocol)
  Debug.Assert m_init
  
  Dim Struct As TStruct: Set Struct = NewTStruct("TApplicationException")
  Proto.WriteStructBegin Struct
  
  Dim Field As TField
  If m_message <> vbNullString Then
    Set Field = NewTField("message", TType_String, 1)
    Proto.WriteFieldBegin Field
    Proto.WriteString m_message
    Proto.WriteFieldEnd
  End If
  
  Set Field = NewTField("type", TType_I32, 2)
  Proto.WriteFieldBegin Field
  Proto.WriteI32 m_type
  Proto.WriteFieldEnd
  
  Proto.WriteFieldStop
  Proto.WriteStructEnd
End Sub

Public Property Get TType() As Byte
  Debug.Assert m_init
  TType = m_type
End Property

Public Property Get Message() As String
  Debug.Assert m_init
  Message = m_message
End Property

' Interface implementation

Private Sub TSerializable_TRead(ByVal Proto As TProtocol)
  TRead Proto
End Sub

Private Sub TSerializable_TWrite(ByVal Proto As TProtocol)
  TWrite Proto
End Sub
