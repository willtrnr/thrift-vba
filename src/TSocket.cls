VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements TTransport

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private m_init As Boolean

Private m_host As String
Private m_port As Long

Private m_socket As LongPtr

Private Sub Class_Initialize()
  m_init = False
  m_host = vbNullString
  m_port = 0
  m_socket = Winsock2.INVALID_SOCKET
End Sub

Public Sub Init(ByVal Host As String, ByVal Port As Long)
  Debug.Assert Not m_init
  m_init = True
  m_host = Host
  m_port = Port
End Sub

Public Function IsOpen() As Boolean
  Debug.Assert m_init
  IsOpen = m_socket <> Winsock2.INVALID_SOCKET
End Function

Public Function Peek() As Boolean
  Debug.Assert m_init
  Peek = IsOpen
End Function

Public Sub TOpen()
  Debug.Assert m_init
  
  If IsOpen Then
    Err.Raise 5, Description:="Socket already connected"
  End If
  
  Winsock2.Startup
  m_socket = Winsock2.Connect(m_host, m_port, Winsock2.AF_UNSPEC, Winsock2.SOCK_STREAM, Winsock2.IPPROTO_TCP)
End Sub

Public Sub TClose()
  Debug.Assert m_init
  If m_socket <> Winsock2.INVALID_SOCKET Then
    Winsock2.Disconnect m_socket
    m_socket = Winsock2.INVALID_SOCKET
  End If
End Sub

Public Function Read(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  
  If Not IsOpen Then
    Err.Raise 5, Description:="Socket is not connected"
  End If
  
  Dim B() As Byte: ReDim B(Length - 1)
  Read = Winsock2.Recv(m_socket, B)
  CopyMemory VarPtr(Buf(Offset)), VarPtr(B(0)), Read
End Function

Public Function ReadAll(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init

  ReadAll = Read(Buf, Offset, Length)
  If ReadAll <> Length Then
    Err.Raise 5, Description:="Cannot read, EOS"
  End If
End Function

Public Sub TWrite(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long)
  Debug.Assert m_init
  
  Dim B() As Byte: ReDim B(Length - 1)
  CopyMemory VarPtr(B(0)), VarPtr(Buf(Offset)), Length
  Winsock2.Send m_socket, B
End Sub

Public Sub Flush()
  Debug.Assert m_init
End Sub


' Interface implementation

Private Function TTransport_IsOpen() As Boolean
  TTransport_IsOpen = IsOpen
End Function

Private Function TTransport_Peek() As Boolean
  TTransport_Peek = Peek
End Function

Private Sub TTransport_TOpen()
  TOpen
End Sub

Private Sub TTransport_TClose()
  TClose
End Sub

Private Function TTransport_Read(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  TTransport_Read = Read(Buf, Offset, Length)
End Function

Private Function TTransport_ReadAll(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  TTransport_ReadAll = ReadAll(Buf, Offset, Length)
End Function

Private Sub TTransport_TWrite(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long)
  TWrite Buf, Offset, Length
End Sub

Private Sub TTransport_Flush()
  Flush
End Sub
