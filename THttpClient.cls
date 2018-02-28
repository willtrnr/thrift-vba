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

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private m_init As Boolean

Private m_url As String
Private m_inputStream As ADODB.Stream
Private m_outputStream As ADODB.Stream

Private Sub Class_Initialize()
  m_init = False
  m_url = vbNullString
  Set m_inputStream = Nothing
  Set m_outputStream = New ADODB.Stream
  m_outputStream.Mode = adModeReadWrite
  m_outputStream.Type = adTypeBinary
  m_outputStream.Open
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
  m_outputStream.Close
  Set m_outputStream = Nothing
End Sub

Public Function Read(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  
  If m_inputStream Is Nothing Then
    Err.Raise 5, Description:="Response buffer is empty, no request"
  End If

  Dim Bytes() As Byte
  Bytes = m_inputStream.Read(Length)
  CopyMemory Buf(Offset), Bytes(0), UBound(Bytes) + 1
  Read = UBound(Bytes) + 1
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
  
  Dim B() As Byte
  ReDim B(Length - 1)
  CopyMemory B(0), Buf(Offset), Length
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

  Xhr.Open "POST", m_url, False
  Xhr.setRequestHeader "Content-Type", "application/x-thrift"
  Xhr.setRequestHeader "Content-Length", m_outputStream.Size
  Xhr.setRequestHeader "Accept", "application/x-thrift"
  Xhr.setRequestHeader "User-Agent", "VBA/THttpClient"
  
  m_outputStream.Position = 0
  Xhr.send m_outputStream.Read(m_outputStream.Size)
  
  m_outputStream.Close
  Set m_outputStream = New ADODB.Stream
  m_outputStream.Mode = adModeReadWrite
  m_outputStream.Type = adTypeBinary
  m_outputStream.Open
  
  If Xhr.Status <> 200 Then
    Err.Raise 5, Description:="HTTP request failed"
  End If
  
  Set m_inputStream = New ADODB.Stream
  m_inputStream.Mode = adModeReadWrite
  m_inputStream.Type = adTypeBinary
  m_inputStream.Open
  m_inputStream.Write Xhr.responseBody
  m_inputStream.Position = 0
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

