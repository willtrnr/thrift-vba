Attribute VB_Name = "ThriftMain"
'@Folder("Thrift")
Option Explicit

Private SeqId As Long

Public Sub Main()
  Dim Trans As TTransport
  Set Trans = TTransportFactory.GetHttpClient("http://127.0.0.1:8888")
  
  Dim Proto As TProtocol
  Set Proto = TProtocolFactory.GetMultiplexedProtocol(TProtocolFactory.GetBinaryProtocol(Trans), "test")

  Debug.Print Add(Proto, 3, 8)
End Sub

Private Function NextSeqId() As Long
  NextSeqId = SeqId
  SeqId = SeqId + 1
End Function

' TODO: This is more or less what the codegen should give

Private Sub Ping(ByVal Proto As TProtocol)
  Dim Message As TMessage
  Dim Struct As TStruct
  Dim Field As TField
  
  Set Message = New TMessage
  Message.Init "ping", TMessageType_Call, NextSeqId
  Proto.WriteMessageBegin Message
  
  Set Struct = New TStruct
  Struct.Init "ping_args"
  Proto.WriteStructBegin Struct
  
  Proto.WriteFieldStop
  Proto.WriteStructEnd
  Proto.WriteMessageEnd
  
  Proto.GetTransport.Flush
  
  Set Message = Proto.ReadMessageBegin
  If Message.TType = TMessageType_Exception Then
    Dim Ex As TApplicationException
    Set Ex = New TApplicationException
    Ex.Read Proto
    Err.Raise 5, Description:=Ex.Message
  End If
  
  Proto.ReadStructBegin
  Do While True
    Set Field = Proto.ReadFieldBegin
    If Field.TType = TType_Stop Then
      Exit Do
    Else
      TProtocolUtil.Skip Proto, Field.TType
    End If
    Proto.ReadFieldEnd
  Loop
  Proto.ReadStructEnd
  Proto.ReadMessageEnd
End Sub

Private Function Add(ByVal Proto As TProtocol, ByVal Num1 As Long, ByVal Num2 As Long) As Long
  Dim Message As TMessage
  Dim Struct As TStruct
  Dim Field As TField
  
  Set Message = New TMessage
  Message.Init "add", TMessageType_Call, NextSeqId
  Proto.WriteMessageBegin Message
  
  Set Struct = New TStruct
  Struct.Init "add_args"
  Proto.WriteStructBegin Struct
  
  Set Field = New TField
  Field.Init "num1", TType_I32, 1
  Proto.WriteFieldBegin Field
  Proto.WriteI32 Num1
  Proto.WriteFieldEnd
  
  Set Field = New TField
  Field.Init "num2", TType_I32, 2
  Proto.WriteFieldBegin Field
  Proto.WriteI32 Num2
  Proto.WriteFieldEnd
  
  Proto.WriteFieldStop
  Proto.WriteStructEnd
  Proto.WriteMessageEnd
  
  Proto.GetTransport.Flush
  
  Set Message = Proto.ReadMessageBegin
  If Message.TType = TMessageType_Exception Then
    Dim Ex As TApplicationException
    Set Ex = New TApplicationException
    Ex.Read Proto
    Err.Raise 5, Description:=Ex.Message
  End If
  
  Proto.ReadStructBegin
  Do While True
    Set Field = Proto.ReadFieldBegin
    If Field.TType = TType_Stop Then
      Exit Do
    ElseIf Field.Id = 0 And Field.TType = TType_I32 Then
      Add = Proto.ReadI32
    Else
      TProtocolUtil.Skip Proto, Field.TType
    End If
    Proto.ReadFieldEnd
  Loop
  Proto.ReadStructEnd
  Proto.ReadMessageEnd
End Function
