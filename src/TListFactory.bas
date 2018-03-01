Attribute VB_Name = "TListFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Public Function NewTList(ByVal ElemType As Byte, ByVal Size As Long) As TList
  Set NewTList = New TList
  NewTList.Init ElemType, Size
End Function
