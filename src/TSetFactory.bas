Attribute VB_Name = "TSetFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Public Function NewTSet(ByVal ElemType As Byte, ByVal Size As Long) As TSet
  Set NewTSet = New TSet
  NewTSet.Init ElemType, Size
End Function
