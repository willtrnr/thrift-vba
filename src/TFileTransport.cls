VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TFileTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Transport")
Option Explicit

Implements TTransport

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private m_init As Boolean

Private m_path As String
Private m_stream As ADODB.Stream

Private Sub Class_Initialize()
  m_init = False
  m_path = vbNullString
  Set m_stream = Nothing
End Sub

Public Sub Init(ByVal Path As String)
  Debug.Assert Not m_init
  m_init = True
  m_path = Path
End Sub

Public Function IsOpen() As Boolean
  Debug.Assert m_init
  IsOpen = Not m_stream Is Nothing
End Function

Public Function Peek() As Boolean
  Debug.Assert m_init
  If Not m_stream Is Nothing Then
    Peek = Not m_stream.EOS
  Else
    Peek = False
  End If
End Function

Public Sub TOpen()
  Debug.Assert m_init
  If m_stream Is Nothing Then
    Set m_stream = New ADODB.Stream
    m_stream.Open
    m_stream.Type = adTypeBinary
    m_stream.LoadFromFile m_path
  End If
End Sub

Public Sub TClose()
  Debug.Assert m_init
  If Not m_stream Is Nothing Then
    m_stream.Close
    Set m_stream = Nothing
  End If
End Sub

Public Function Read(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  
  If Not IsOpen Then
    Err.Raise 5, Description:="Must open before reading"
  End If
  
  Dim B As Variant
  B = m_stream.Read(Length)
  CopyMemory VarPtr(Buf(Offset)), VarPtr(B(0)), UBound(B) + 1
  Read = UBound(B) + 1
End Function

Public Function ReadAll(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  
  If Not IsOpen Then
    Err.Raise 5, Description:="Must open before reading"
  End If
  
  Dim B As Variant
  B = m_stream.Read(Length)
  If UBound(B) + 1 <> Length Then
    Err.Raise 5, Description:="End of file reached"
  End If
  CopyMemory VarPtr(Buf(Offset)), VarPtr(B(0)), Length
  ReadAll = Length
End Function

Public Sub TWrite(ByRef Buf() As Byte, ByVal Offset As Long, ByVal Length As Long)
  Debug.Assert m_init
  Err.Raise 5, Description:="Not supported"
End Sub

Public Sub Flush()
  Debug.Assert m_init
  Err.Raise 5, Description:="Not supported"
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
