Attribute VB_Name = "TTransportFactory"
'@Folder("Thrift.Transport")
Option Explicit

Public Function NewTFileTransport(ByVal Path As String) As TFileTransport
  Set NewTFileTransport = New TFileTransport
  NewTFileTransport.Init Path
End Function

Public Function NewTHttpClient(ByVal Url As String) As THttpClient
  Set NewTHttpClient = New THttpClient
  NewTHttpClient.Init Url
End Function
