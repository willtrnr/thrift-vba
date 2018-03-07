VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "THttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Transport")
Option Explicit

Implements TTransport

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private m_init As Boolean

Private m_url As String
Private m_inputStream As ADODB.Stream
Private m_outputStream As ADODB.Stream

Private Sub Class_Initialize()
  m_init = False
  m_url = vbNullString
  Set m_inputStream = Nothing
  Set m_outputStream = New ADODB.Stream
  With m_outputStream
    .Mode = adModeReadWrite
    .Type = adTypeBinary
    .Open
  End With
End Sub

Public Sub Init(ByVal Url As String)
  Debug.Assert Not m_init
  m_init = True
  m_url = Url
End Sub

Public Function IsOpen() As Boolean
  Debug.Assert m_init
  IsOpen = True
End Function

Public Function Peek() As Boolean
  Debug.Assert m_init
  Peek = True
End Function

Public Sub TOpen()
  Debug.Assert m_init
End Sub

Public Sub TClose()
  Debug.Assert m_init
  If Not m_inputStream Is Nothing Then
    m_inputStream.Close
    Set m_inputStream = Nothing
  End If
  If Not m_outputStream Is Nothing Then
    m_outputStream.Close
    Set m_outputStream = Nothing
  End If
End Sub

Public Function Read(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  
  If m_inputStream Is Nothing Then
    Err.Raise 5, Description:="Response Buffer is empty, no request"
  End If

  Dim Bytes() As Byte: Bytes = m_inputStream.Read(Length)
  CopyMemory VarPtr(Buffer(Offset)), VarPtr(Bytes(0)), UBound(Bytes) + 1
  Read = UBound(Bytes) + 1
End Function

Public Function ReadAll(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  
  ReadAll = Read(Buffer, Offset, Length)
  If ReadAll <> Length Then
    Err.Raise 5, Description:="Cannot read, EOS"
  End If
End Function

Public Sub TWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long)
  Debug.Assert m_init
  
  Dim B() As Byte: ReDim B(Length - 1)
  CopyMemory VarPtr(B(0)), VarPtr(Buffer(Offset)), Length
  m_outputStream.Write B
End Sub

Public Sub Flush()
  Debug.Assert m_init
  
  If Not m_inputStream Is Nothing Then
    m_inputStream.Close
    Set m_inputStream = Nothing
  End If
  
  Dim Xhr As MSXML2.ServerXMLHTTP60
  Set Xhr = New MSXML2.ServerXMLHTTP60
  With Xhr
    .Open "POST", m_url, False
    .setRequestHeader "Content-Type", "application/x-thrift"
    .setRequestHeader "Content-Length", m_outputStream.Size
    .setRequestHeader "Accept", "application/x-thrift"
    .setRequestHeader "User-Agent", "VBA/THttpClient"
  End With
  
  m_outputStream.Position = 0
  Xhr.Send m_outputStream.Read(m_outputStream.Size)
  m_outputStream.Close

  Set m_outputStream = New ADODB.Stream
  With m_outputStream
    .Mode = adModeReadWrite
    .Type = adTypeBinary
    .Open
  End With
  
  If Xhr.Status <> 200 Then
    Err.Raise 5, Description:="HTTP request failed"
  End If
  
  Set m_inputStream = New ADODB.Stream
  With m_inputStream
    .Mode = adModeReadWrite
    .Type = adTypeBinary
    .Open
    .Write Xhr.responseBody
    .Position = 0
  End With
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

Private Function TTransport_Read(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  TTransport_Read = Read(Buffer, Offset, Length)
End Function

Private Function TTransport_ReadAll(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  TTransport_ReadAll = ReadAll(Buffer, Offset, Length)
End Function

Private Sub TTransport_TWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long)
  TWrite Buffer, Offset, Length
End Sub

Private Sub TTransport_Flush()
  Flush
End Sub

