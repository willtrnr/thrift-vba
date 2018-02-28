VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Protocol")
Option Explicit

Private m_init As Boolean

Private m_name As String

Private Sub Class_Initialize()
  m_init = False
  m_name = vbNullString
End Sub

Public Sub Init(ByVal N As String)
  Debug.Assert Not m_init
  m_init = True
  m_name = N
End Sub

Public Property Get Name() As String
  Debug.Assert m_init
  Name = m_name
End Property
