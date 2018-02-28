Attribute VB_Name = "TProtocolFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Public Function GetBinaryProtocol(ByVal Trans As TTransport, Optional ByVal StrictRead As Boolean = False, Optional ByVal StrictWrite As Boolean = True) As TProtocol
  Dim P As TBinaryProtocol
  Set P = New TBinaryProtocol
  P.Init Trans, StrictRead, StrictWrite
  Set GetBinaryProtocol = P
End Function

Public Function GetMultiplexedProtocol(ByVal Proto As TProtocol, ByVal ServiceName As String) As TProtocol
  Dim P As TMultiplexedProtocol
  Set P = New TMultiplexedProtocol
  P.Init Proto, ServiceName
  Set GetMultiplexedProtocol = P
End Function
