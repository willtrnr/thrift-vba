Attribute VB_Name = "TMapFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Public Function NewTMap(ByVal KeyType As Byte, ByVal ValueType As Byte, ByVal Size As Long) As TMap
  Set NewTMap = New TMap
  NewTMap.Init KeyType, ValueType, Size
End Function
