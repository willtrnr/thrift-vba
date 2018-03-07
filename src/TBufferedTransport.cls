VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TBufferedTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Transport")
Option Explicit

Implements TTransport

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private m_init As Boolean

Private m_bufferSize As Long
Private m_inputBuffer As ADODB.Stream
Private m_outputBuffer As ADODB.Stream
Private m_trans As TTransport

Private Sub Class_Initialize()
  m_init = False
  m_bufferSize = 0
  Set m_inputBuffer = Nothing
  Set m_outputBuffer = Nothing
  Set m_trans = Nothing
End Sub

Public Sub Init(ByVal Trans As TTransport, Optional ByVal BufferSize As Long = 1024)
  Debug.Assert Not m_init
  If BufferSize <= 0 Then
    Err.Raise 5, Description:="Buffer size must be larger than 0"
  End If
  m_init = True
  m_bufferSize = 1024
  Set m_inputBuffer = Nothing
  Set m_outputBuffer = Nothing
  Set m_trans = Trans
End Sub

Public Function IsOpen() As Boolean
  Debug.Assert m_init
  IsOpen = (Not m_inputBuffer Is Nothing) And (Not m_outputBuffer Is Nothing) And m_trans.IsOpen
End Function

Public Function Peek() As Boolean
  Debug.Assert m_init
  Peek = (Not m_inputBuffer Is Nothing) And (Not m_outputBuffer Is Nothing) And m_trans.Peek
End Function

Public Sub TOpen()
  Debug.Assert m_init
  Set m_inputBuffer = New ADODB.Stream
  With m_inputBuffer
    .Mode = adModeReadWrite
    .Type = adTypeBinary
    .Open
  End With
  Set m_outputBuffer = New ADODB.Stream
  With m_outputBuffer
    .Mode = adModeReadWrite
    .Type = adTypeBinary
    .Open
  End With
  m_trans.TOpen
End Sub

Public Sub TClose()
  Debug.Assert m_init
  If Not m_inputBuffer Is Nothing Then
    m_inputBuffer.Close
    Set m_outputBuffer = Nothing
  End If
  If Not m_outputBuffer Is Nothing Then
    m_outputBuffer.Close
    Set m_outputBuffer = Nothing
  End If
  m_trans.TClose
End Sub

Public Function Read(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  CheckOpen

  Dim B() As Byte
  If m_inputBuffer.EOS Then
    m_inputBuffer.Close
    Set m_inputBuffer = New ADODB.Stream
    With m_inputBuffer
      .Mode = adModeReadWrite
      .Type = adTypeBinary
      .Open
    End With

    ReDim B(m_bufferSize - 1)
    Dim Filled As Long: Filled = m_trans.Read(B, 0, m_bufferSize)
    ReDim Preserve B(Filled - 1)

    With m_inputBuffer
      .Write B
      .Position = 0
    End With
  End If

  B = m_inputBuffer.Read(Length)
  CopyMemory VarPtr(Buffer(Offset)), VarPtr(B(0)), UBound(B) + 1
  Read = UBound(B) + 1
End Function

Public Function ReadAll(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init
  CheckOpen

  ReadAll = 0
  While ReadAll < Length
    Dim Res As Long
    Res = Read(Buffer, Offset + ReadAll, Length - ReadAll)
    If Res <= 0 Then
      Err.Raise 5, Description:="Cannot read EOS"
    End If
    ReadAll = ReadAll + Res
  Wend
End Function

Public Sub TWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long)
  Debug.Assert m_init
  CheckOpen
  
  Dim B() As Byte
  If m_bufferSize - m_outputBuffer.Position >= Length Then
    ReDim B(Length - 1)
    CopyMemory VarPtr(B(0)), VarPtr(Buffer(Offset)), Length
    m_outputBuffer.Write B
  Else
    Dim Wrote As Long: Wrote = 0
    If m_outputBuffer.Size > 0 And m_outputBuffer.Position < m_bufferSize Then
      ReDim B(m_bufferSize - m_outputBuffer.Size - 1)
      CopyMemory VarPtr(B(0)), VarPtr(Buffer(Offset)), UBound(B) + 1
      m_outputBuffer.Write B

      m_outputBuffer.Position = 0
      B = m_outputBuffer.Read(m_outputBuffer.Size)
      m_trans.TWrite B, 0, UBound(B) + 1
      Wrote = Wrote + UBound(B) + 1

      m_outputBuffer.Close
      Set m_outputBuffer = New ADODB.Stream
      With m_outputBuffer
        .Mode = adModeReadWrite
        .Type = adTypeBinary
        .Open
      End With
    End If
    While Length - Wrote >= m_bufferSize
      m_trans.TWrite Buffer, Offset + Wrote, m_bufferSize
      Wrote = Wrote + m_bufferSize
    Wend
    If Length - Wrote > 0 Then
      ReDim B(Length - Wrote - 1)
      CopyMemory VarPtr(B(0)), VarPtr(Buffer(Offset + Wrote)), Length - Wrote
      m_outputBuffer.Write B
    End If
  End If
End Sub

Public Sub Flush()
  Debug.Assert m_init
  CheckOpen
  If m_outputBuffer.Size > 0 Then
    m_outputBuffer.Position = 0
    Dim Buffer() As Byte: Buffer = m_outputBuffer.Read(m_outputBuffer.Size)
    m_trans.TWrite Buffer, 0, UBound(Buffer) + 1
    
    m_outputBuffer.Close
    Set m_outputBuffer = New ADODB.Stream
    With m_outputBuffer
      .Mode = adModeReadWrite
      .Type = adTypeBinary
      .Open
    End With
  End If
  m_trans.Flush
End Sub

Private Sub CheckOpen()
  Debug.Assert m_init
  If Not IsOpen Then
    Err.Raise 5, Description:="Transport is not open"
  End If
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
