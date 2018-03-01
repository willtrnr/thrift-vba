Attribute VB_Name = "TFieldFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Public Function NewTField(ByVal Name As String, ByVal TType As Byte, ByVal Id As Integer) As TField
  Set NewTField = New TField
  NewTField.Init Name, TType, Id
End Function
