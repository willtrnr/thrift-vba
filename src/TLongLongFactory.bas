Attribute VB_Name = "TLongLongFactory"
Option Explicit

Private TLongLong0 As TLongLong
Private TLongLong1 As TLongLong
Private TLongLongN1 As TLongLong

Public Function NewTLongLong(ByVal Value As Variant) As TLongLong
  If Value = 0 Then
    If TLongLong0 Is Nothing Then
      Set TLongLong0 = New TLongLong
      TLongLong0.Init 0
    End If
    Set NewTLongLong = TLongLong0
  ElseIf Value = 1 Then
    If TLongLong1 Is Nothing Then
      Set TLongLong1 = New TLongLong
      TLongLong1.Init 1
    End If
    Set NewTLongLong = TLongLong1
  ElseIf Value = -1 Then
    If TLongLongN1 Is Nothing Then
      Set TLongLongN1 = New TLongLong
      TLongLongN1.Init -1
    End If
    Set NewTLongLong = TLongLongN1
  Else
    Set NewTLongLong = New TLongLong
    NewTLongLong.Init Value
  End If
End Function
