Attribute VB_Name = "TStructFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Private AnonymousStruct As TStruct

Public Function NewTStruct(ByVal Name As String) As TStruct
  If Name = vbNullString Then
    If AnonymousStruct Is Nothing Then
      Set AnonymousStruct = New TStruct
      AnonymousStruct.Init vbNullString
    End If
    Set NewTStruct = AnonymousStruct
  Else
    Set NewTStruct = New TStruct
    NewTStruct.Init Name
  End If
End Function
