Attribute VB_Name = "TMessageFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Public Function NewTMessage(ByVal Name As String, ByVal TType As Byte, ByVal SeqId As Long) As TMessage
  Set NewTMessage = New TMessage
  NewTMessage.Init Name, TType, SeqId
End Function
