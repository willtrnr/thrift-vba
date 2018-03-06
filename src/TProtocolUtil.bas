Attribute VB_Name = "TProtocolUtil"
'@Folder("Thrift.Protocol")
Option Explicit

Private Const MAX_SKIP_DEPTH As Long = &H7FFFFFFF

Public Sub Skip(ByVal Proto As TProtocol, ByVal TType As Byte, Optional ByVal MaxDepth As Long = MAX_SKIP_DEPTH)
  If MaxDepth <= 0 Then
    Err.Raise 5, Description:="Maximum skip depth exceeded"
  End If

  Dim I As Long
  Dim Field As TField
  Select Case TType
    Case TType_Bool
      Proto.ReadBool
    Case TType_Byte
      Proto.ReadByte
    Case TType_Double
      Proto.ReadDouble
    Case TType_I16
      Proto.ReadI16
    Case TType_I32
      Proto.ReadI32
    Case TType_String
      Proto.ReadString
    Case TType_Struct
      Proto.ReadStructBegin
      Do While True
        Set Field = Proto.ReadFieldBegin
        If Field.TType = TType_Stop Then
          Exit Do
        End If
        Skip Proto, Field.TType, MaxDepth - 1
        Proto.ReadFieldEnd
      Loop
      Proto.ReadStructEnd
    Case TType_Map
      Dim Map As TMap
      Set Map = Proto.ReadMapBegin
      For I = 1 To Map.Size
        Skip Proto, Map.KeyType, MaxDepth - 1
        Skip Proto, Map.ValueType, MaxDepth - 1
      Next I
      Proto.ReadMapEnd
    Case TType_Set
      Dim Set_ As TSet
      Set Set_ = Proto.ReadSetBegin
      For I = 1 To Set_.Size
        Skip Proto, Set_.ElemType, MaxDepth - 1
      Next I
      Proto.ReadSetEnd
    Case TType_List
      Dim List As TList
      Set List = Proto.ReadListBegin
      For I = 1 To List.Size
        Skip Proto, List.ElemType, MaxDepth - 1
      Next I
      Proto.ReadListEnd
    Case Else
      Err.Raise 5, Description:="Unrecognized type"
  End Select
End Sub
