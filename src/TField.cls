VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Protocol")
Option Explicit

Private m_init As Boolean

Private m_name As String
Private m_type As Byte
Private m_id As Integer

Private Sub Class_Initialize()
  m_init = False
  m_name = vbNullString
  m_type = 0
  m_id = 0
End Sub

Public Sub Init(ByVal Name As String, ByVal TType As Byte, ByVal Id As Integer)
  Debug.Assert Not m_init
  m_init = True
  m_name = Name
  m_type = TType
  m_id = Id
End Sub

Public Property Get Name() As String
  Debug.Assert m_init
  Name = m_name
End Property

Public Property Get TType() As Byte
  Debug.Assert m_init
  TType = m_type
End Property

Public Property Get Id() As Integer
  Debug.Assert m_init
  Id = m_id
End Property
