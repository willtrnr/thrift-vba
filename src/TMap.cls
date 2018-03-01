VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Protocol")
Option Explicit

Private m_init As Boolean

Private m_keyType As Byte
Private m_valueType As Byte
Private m_size As Long

Private Sub Class_Initialize()
  m_init = False
  m_keyType = 0
  m_valueType = 0
  m_size = 0
End Sub

Public Sub Init(ByVal KeyType As Byte, ByVal ValueType As Byte, ByVal Size As Long)
  Debug.Assert Not m_init
  m_init = True
  m_keyType = KeyType
  m_valueType = ValueType
  m_size = Size
End Sub

Public Property Get KeyType() As Byte
  Debug.Assert m_init
  KeyType = m_keyType
End Property

Public Property Get ValueType() As Byte
  Debug.Assert m_init
  ValueType = m_valueType
End Property

Public Property Get Size() As Long
  Debug.Assert m_init
  Size = m_size
End Property
