Attribute VB_Name = "TSocketFactory"
Option Explicit

Public Function NewTSocket(ByVal Host As String, ByVal Port As Long) As TSocket
  Set NewTSocket = New TSocket
  NewTSocket.Init Host, Port
End Function
