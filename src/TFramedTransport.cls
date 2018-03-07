VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TFramedTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Transport")
Option Explicit

Implements TTransport

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private m_init As Boolean

Private m_maxLength As Long
Private m_inputBuffer As ADODB.Stream
Private m_outputBuffer As ADODB.Stream
Private m_trans As TTransport

Private Sub Class_Initialize()
  m_init = False
  m_maxLength = 0
  Set m_inputBuffer = Nothing
  Set m_outputBuffer = Nothing
  Set m_trans = Nothing
End Sub

Public Sub Init(ByVal Trans As TTransport, Optional ByVal MaxLength As Long = 16& * 1024& * 1024&)
  Debug.Assert Not m_init
  If MaxLength <= 0 Then
    Err.Raise 5, Description:="Maximum length must be larger than 0"
  End If
  m_init = True
  m_maxLength = MaxLength
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
  
  If m_inputBuffer.EOS Then
    Dim SizeBuffer(3) As Byte
    m_trans.ReadAll SizeBuffer, 0, 4
    
    Dim Size As Long: Size = DecodeFrameSize(SizeBuffer)
    If Size < 0 Then
      Err.Raise 5, Description:="Read a negative frame size"
    ElseIf Size > m_maxLength Then
      Err.Raise 5, Description:="Frame size larger than max length"
    End If
    
    Dim FrameBuffer() As Byte: ReDim FrameBuffer(Size - 1)
    m_trans.ReadAll FrameBuffer, 0, Size

    m_inputBuffer.Close
    Set m_inputBuffer = New ADODB.Stream
    With m_inputBuffer
      .Mode = adModeReadWrite
      .Type = adTypeBinary
      .Open
      .Write FrameBuffer
      .Position = 0
    End With
  End If
  
  Dim B() As Byte: B = m_inputBuffer.Read(Length)
  CopyMemory VarPtr(Buffer(Offset)), VarPtr(B(0)), UBound(B) + 1
  Read = UBound(B) + 1
End Function

Public Function ReadAll(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long) As Long
  Debug.Assert m_init

  ReadAll = 0
  While ReadAll < Length
    Dim Res As Long
    Res = Read(Buffer, Offset + ReadAll, Length - ReadAll)
    If Res <= 0 Then
      Err.Raise 5, Description:="Cannot read, EOS"
    End If
    ReadAll = ReadAll + Res
  Wend
End Function

Public Sub TWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Length As Long)
  Debug.Assert m_init
  CheckOpen
  
  Dim B() As Byte: ReDim B(Length - 1)
  CopyMemory VarPtr(B(0)), VarPtr(Buffer(Offset)), Length
  m_outputBuffer.Write B
End Sub

Public Sub Flush()
  Debug.Assert m_init
  CheckOpen
  
  If m_outputBuffer.Size > 0 Then
    m_outputBuffer.Position = 0
    Dim Buffer() As Byte: Buffer = m_outputBuffer.Read(m_outputBuffer.Size)
    
    Dim SizeBuffer(3) As Byte
    EncodeFrameSize UBound(Buffer) + 1, SizeBuffer
    m_trans.TWrite SizeBuffer, 0, 4
  
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

Private Sub EncodeFrameSize(ByVal Size As Long, ByRef Buffer() As Byte)
  Debug.Assert UBound(Buffer) >= 3
  Buffer(0) = CByte((Size \ &H1000000) And &HFF&)
  Buffer(1) = CByte((Size \ &H10000) And &HFF&)
  Buffer(2) = CByte((Size \ &H100&) And &HFF&)
  Buffer(3) = CByte(Size And &HFF&)
End Sub

Private Function DecodeFrameSize(ByRef Buffer() As Byte) As Long
  Debug.Assert UBound(Buffer) >= 3
  DecodeFrameSize = _
    CLng(Buffer(0) And &H7F) * &H1000000 + _
    CLng(Buffer(1)) * &H10000 + _
    CLng(Buffer(2)) * &H100& + _
    CLng(Buffer(3))
  If (Buffer(0) And &H80) = &H80 Then
    DecodeFrameSize = DecodeFrameSize Or &H80
  End If
End Function


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

