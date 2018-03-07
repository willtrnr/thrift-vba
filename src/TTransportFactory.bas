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

Public Function NewTSocket(ByVal Host As String, ByVal Port As Long) As TSocket
  Set NewTSocket = New TSocket
  NewTSocket.Init Host, Port
End Function

Public Function NewTBufferedTransport(ByVal Trans As TTransport, Optional ByVal BufferSize As Long = 1024) As TBufferedTransport
  Set NewTBufferedTransport = New TBufferedTransport
  NewTBufferedTransport.Init Trans, BufferSize
End Function
