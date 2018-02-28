Attribute VB_Name = "TTransportFactory"
'@Folder("Thrift.Transport")
Option Explicit

Public Function GetFileTransport(ByVal Path As String) As TFileTransport
  Dim T As TFileTransport
  Set T = New TFileTransport
  T.Init Path
  Set GetFileTransport = T
End Function

Public Function GetHttpClient(ByVal Url As String) As THttpClient
  Dim T As THttpClient
  Set T = New THttpClient
  T.Init Url
  Set GetHttpClient = T
End Function
