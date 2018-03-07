Attribute VB_Name = "Winsock2"
'@Folder("Win32")
Option Explicit

Public Const AF_UNSPEC As Long = 0
Public Const AF_INET As Long = 2
Public Const AF_NETBIOS As Long = 17
Public Const AF_INET6 As Long = 23
Public Const AF_IRDA As Long = 26
Public Const AF_BTH As Long = 32

Public Const SOCK_STREAM As Long = 1
Public Const SOCK_DGRAM As Long = 2
Public Const SOCK_RAW As Long = 3
Public Const SOCK_RDM As Long = 4
Public Const SOCK_SEQPACKET As Long = 5

Public Const IPPROTO_TCP As Long = 6
Public Const IPPROTO_UDP As Long = 17
Public Const IPPROTO_RM As Long = 113

Public Const SD_RECEIVE As Long = 0
Public Const SD_SEND As Long = 1
Public Const SD_BOTH As Long = 2

Public Const INVALID_SOCKET As LongPtr = -1
Public Const SOCKET_ERROR As Long = -1

Private Const WSADESCRIPTION_LEN As Long = 256
Private Const WSASYS_STATUS_LEN As Long = 128

Private Type WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription(WSADESCRIPTION_LEN) As Byte
  szSystemStatus(WSASYS_STATUS_LEN) As Byte
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpVendorInfo As String
End Type

Private Type sockaddr
  sa_family As Integer
  sa_data(14) As Byte
End Type

Private Type addrinfo
  ai_flags As Long
  ai_family As Long
  ai_socktype As Long
  ai_protocol As Long
  ai_addrlen As Long
  ai_canonname As LongPtr
  ai_addr As LongPtr
  ai_next As LongPtr
End Type

Private Declare PtrSafe Function WS2_WSAStartup Lib "Ws2_32.dll" Alias "WSAStartup" (ByVal wVersionRequested As Long, ByRef lpWSAData As WSADATA) As Long
Private Declare PtrSafe Function WS2_WSACleanup Lib "Ws2_32.dll" Alias "WSACleanup" () As Long
Private Declare PtrSafe Function WS2_WSAGetLastError Lib "Ws2_32.dll" Alias "WSAGetLastError" () As Long
Private Declare PtrSafe Function WS2_getaddrinfo Lib "Ws2_32.dll" Alias "getaddrinfo" (ByVal pNodeName As String, ByVal pServiceName As String, ByRef pHints As addrinfo, ByVal ppResult As LongPtr) As Long
Private Declare PtrSafe Sub WS2_freeaddrinfo Lib "Ws2_32.dll" Alias "freeaddrinfo" (ByVal ai As LongPtr)
Private Declare PtrSafe Function WS2_socket Lib "Ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal SockType As Long, ByVal Protocol As Long) As LongPtr
Private Declare PtrSafe Function WS2_connect Lib "Ws2_32.dll" Alias "connect" (ByVal s As LongPtr, ByRef Name As sockaddr, ByVal namelen As Long) As Long
Private Declare PtrSafe Function WS2_send Lib "Ws2_32.dll" Alias "send" (ByVal s As LongPtr, ByVal buf As LongPtr, ByVal len_ As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function WS2_recv Lib "Ws2_32.dll" Alias "recv" (ByVal s As LongPtr, ByVal buf As LongPtr, ByVal len_ As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function WS2_shutdown Lib "Ws2_32.dll" Alias "shutdown" (ByVal s As LongPtr, ByVal how As Long) As Long
Private Declare PtrSafe Function WS2_closesocket Lib "Ws2_32.dll" Alias "closesocket" (ByVal s As LongPtr) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private OpenCount As Long

Private Function MAKEWORD(ByVal bLow As Byte, ByVal bHigh As Byte) As Integer
  MAKEWORD = CInt((CLng(bHigh) * &H100&) Or CLng(bLow))
End Function

Public Function Connect(ByVal Host As String, ByVal Port As Long, ByVal Family As Long, ByVal SockType As Long, ByVal Protocol As Long) As LongPtr
  If OpenCount <= 0 Then
    Dim Data As WSADATA
    Dim Status As Long: Status = WS2_WSAStartup(MAKEWORD(2, 2), Data)
    Debug.Assert Status = 0
  End If

  Dim Result As Long

  Dim Hints As addrinfo
  Hints.ai_family = Family
  Hints.ai_socktype = SockType
  Hints.ai_protocol = Protocol
  
  Dim InfoPtr As LongPtr
  Result = WS2_getaddrinfo(Host, CStr(Port), Hints, VarPtr(InfoPtr))
  If Result <> 0 Then
    Err.Raise WS2_WSAGetLastError(), Description:="Invalid host or port"
  End If
  
  Dim Info As addrinfo
  CopyMemory VarPtr(Info), InfoPtr, LenB(Info)

  Dim Sock As LongPtr: Sock = WS2_socket(Info.ai_family, Info.ai_socktype, Info.ai_protocol)
  If Sock = INVALID_SOCKET Then
      WS2_freeaddrinfo InfoPtr
    Err.Raise WS2_WSAGetLastError(), Description:="Invalid socket"
  End If

  Dim Addr As sockaddr
  CopyMemory VarPtr(Addr), Info.ai_addr, Info.ai_addrlen

  WS2_freeaddrinfo InfoPtr

  Result = WS2_connect(Sock, Addr, Info.ai_addrlen)
  If Result = SOCKET_ERROR Then
    WS2_closesocket Sock
    Err.Raise WS2_WSAGetLastError(), Description:="Cannot connect socket"
  End If

  OpenCount = OpenCount + 1
  Connect = Sock
End Function

Public Function Send(ByVal Sock As LongPtr, ByVal Buffer As LongPtr, ByVal Length As Long, Optional ByVal Flags As Long = 0) As Long
  Dim Result As Long: Result = WS2_send(Sock, Buffer, Length, Flags)
  If Result = SOCKET_ERROR Then
    Err.Raise WS2_WSAGetLastError(), Description:="Send error"
  End If
  Send = Result
End Function

Public Function Recv(ByVal Sock As LongPtr, ByVal Buffer As LongPtr, ByVal Length As Long, Optional ByVal Flags As Long = 0) As Long
  Dim Result As Long: Result = WS2_recv(Sock, Buffer, Length, Flags)
  If Result = SOCKET_ERROR Then
    Err.Raise WS2_WSAGetLastError(), Description:="Receive error"
  End If
  Recv = Result
End Function

Public Sub Disconnect(ByVal Sock As LongPtr)
  WS2_shutdown Sock, SD_BOTH
  WS2_closesocket Sock

  OpenCount = OpenCount - 1
  If OpenCount <= 0 Then
    WS2_WSACleanup
  End If
End Sub
