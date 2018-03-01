Attribute VB_Name = "TStructFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Public Function NewTStruct(ByVal Name As String) As TStruct
  Set NewTStruct = New TStruct
  NewTStruct.Init Name
End Function
