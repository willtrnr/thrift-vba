VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TLongLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_init As Boolean

Private m_bytes(7) As Byte

Private Sub Class_Initialize()
  m_init = False
  m_bytes(0) = 0
  m_bytes(1) = 0
  m_bytes(2) = 0
  m_bytes(3) = 0
  m_bytes(4) = 0
  m_bytes(5) = 0
  m_bytes(6) = 0
  m_bytes(7) = 0
End Sub

Public Sub Init(ByVal Value As Variant)
  Debug.Assert Not m_init
  m_init = True
  Select Case TypeName(Value)
    Case "Byte()"
      If LBound(Value) <> 0 Or UBound(Value) <> 7 Then
        Err.Raise 5, Description:="Wrong byte count for TLongLong"
      Else
        m_bytes(0) = Value(0)
        m_bytes(1) = Value(1)
        m_bytes(2) = Value(2)
        m_bytes(3) = Value(3)
        m_bytes(4) = Value(4)
        m_bytes(5) = Value(5)
        m_bytes(6) = Value(6)
        m_bytes(7) = Value(7)
      End If
    Case "Byte"
      m_bytes(7) = CByte(Value)
    Case "Integer"
      m_bytes(6) = CByte((Value And &H7F00) \ &H100)
      m_bytes(7) = CByte(Value And &HFF)
      If Value < 0 Then
        m_bytes(0) = &HFF
        m_bytes(1) = &HFF
        m_bytes(2) = &HFF
        m_bytes(3) = &HFF
        m_bytes(4) = &HFF
        m_bytes(5) = &HFF
        m_bytes(6) = m_bytes(6) Or &H80
      End If
    Case "Long"
      m_bytes(4) = CByte((Value And &H7F000000) \ &H1000000)
      m_bytes(5) = CByte((Value And &HFF0000) \ &H10000)
      m_bytes(6) = CByte((Value And &HFF00&) \ &H100&)
      m_bytes(7) = CByte(Value And &HFF&)
      If Value < 0 Then
        m_bytes(0) = &HFF
        m_bytes(1) = &HFF
        m_bytes(2) = &HFF
        m_bytes(3) = &HFF
        m_bytes(4) = m_bytes(4) Or &H80
      End If
    Case Else
      Err.Raise 5, Description:="Cannot convert to TLongLong"
  End Select
End Sub

Public Property Get Bytes() As Byte()
  Debug.Assert m_init
  Dim Tmp(7) As Byte
  Tmp(0) = m_bytes(0)
  Tmp(1) = m_bytes(1)
  Tmp(2) = m_bytes(2)
  Tmp(3) = m_bytes(3)
  Tmp(4) = m_bytes(4)
  Tmp(5) = m_bytes(5)
  Tmp(6) = m_bytes(6)
  Tmp(7) = m_bytes(7)
  Bytes = Tmp
End Property

Public Property Get IsValidLong() As Boolean
  Debug.Assert m_init
  IsValidLong = False
  If m_bytes(0) = 0 And m_bytes(1) = 0 And m_bytes(2) = 0 And m_bytes(3) = 0 And (m_bytes(4) And &H80) = 0 Then
    IsValidLong = True
  ElseIf m_bytes(0) = 255 And m_bytes(1) = 255 And m_bytes(2) = 255 And m_bytes(3) = 255 And (m_bytes(4) And &H80) = &H80 Then
    IsValidLong = True
  End If
End Property

Public Property Get AsLong() As Long
  Debug.Assert m_init
  If IsValidLong Then
    AsLong = CLng(m_bytes(4) And &H7F) * &H1000000 + CLng(m_bytes(5)) * &H10000 + CLng(m_bytes(6)) * &H100& + CLng(m_bytes(7))
    If m_bytes(4) And &H80 = &H80 Then
      AsLong = AsLong Or &H80000000
    End If
  Else
    Err.Raise 6, Description:="Overflow"
  End If
End Property

Public Property Get AsDouble() As Double
  If (m_bytes(0) And &H80) = &H80 Then
    AsDouble = CDbl(Not m_bytes(0))
    AsDouble = AsDouble * &H100 + CDbl(Not m_bytes(1))
    AsDouble = AsDouble * &H100 + CDbl(Not m_bytes(2))
    AsDouble = AsDouble * &H100 + CDbl(Not m_bytes(3))
    AsDouble = AsDouble * &H100 + CDbl(Not m_bytes(4))
    AsDouble = AsDouble * &H100 + CDbl(Not m_bytes(5))
    AsDouble = AsDouble * &H100 + CDbl(Not m_bytes(6))
    AsDouble = AsDouble * &H100 + CDbl(Not m_bytes(7))
    AsDouble = -(AsDouble + 1)
  Else
    AsDouble = CDbl(m_bytes(0))
    AsDouble = AsDouble * &H100 + CDbl(m_bytes(1))
    AsDouble = AsDouble * &H100 + CDbl(m_bytes(2))
    AsDouble = AsDouble * &H100 + CDbl(m_bytes(3))
    AsDouble = AsDouble * &H100 + CDbl(m_bytes(4))
    AsDouble = AsDouble * &H100 + CDbl(m_bytes(5))
    AsDouble = AsDouble * &H100 + CDbl(m_bytes(6))
    AsDouble = AsDouble * &H100 + CDbl(m_bytes(7))
  End If
End Property

Public Property Get AsString() As String
  ' TODO
  AsString = CStr(AsDouble)
End Property

Public Function Equals(ByVal Other As TLongLong) As Boolean
  Dim OtherBytes() As Byte
  OtherBytes = Other.Bytes
  Equals = _
    m_bytes(0) = OtherBytes(0) And _
    m_bytes(1) = OtherBytes(1) And _
    m_bytes(2) = OtherBytes(2) And _
    m_bytes(3) = OtherBytes(3) And _
    m_bytes(4) = OtherBytes(4) And _
    m_bytes(5) = OtherBytes(5) And _
    m_bytes(6) = OtherBytes(6) And _
    m_bytes(7) = OtherBytes(7)
End Function
