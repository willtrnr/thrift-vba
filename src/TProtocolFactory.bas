Attribute VB_Name = "TProtocolFactory"
'@Folder("Thrift.Protocol")
Option Explicit

Private Const NO_LENGTH_LIMIT As Long = -1

Public Function NewTBinaryProtocol(ByVal Trans As TTransport, Optional ByVal StrictRead As Boolean = False, Optional ByVal StrictWrite As Boolean = True, Optional ByVal StringLengthLimit As Long = NO_LENGTH_LIMIT, Optional ByVal ContainerLengthLimit As Long = NO_LENGTH_LIMIT) As TBinaryProtocol
  Set NewTBinaryProtocol = New TBinaryProtocol
  NewTBinaryProtocol.Init Trans, StrictRead, StrictWrite, StringLengthLimit, ContainerLengthLimit
End Function

Public Function NewTMultiplexedProtocol(ByVal Proto As TProtocol, ByVal ServiceName As String) As TMultiplexedProtocol
  Set NewTMultiplexedProtocol = New TMultiplexedProtocol
  NewTMultiplexedProtocol.Init Proto, ServiceName
End Function
