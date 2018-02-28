Attribute VB_Name = "TMessageType"
'@Folder("Thrift.Protocol")
Option Explicit

Public Const TMessageType_Call As Byte = 1
Public Const TMessageType_Reply As Byte = 2
Public Const TMessageType_Exception As Byte = 3
Public Const TMessageType_OneWay As Byte = 4
