VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Thrift.Protocol")
Option Explicit

Private m_init As Boolean

Private m_elemType As Byte
Private m_size As Long

Private Sub Class_Initialize()
  m_init = False
  m_elemType = 0
  m_size = 0
End Sub

Public Sub Init(ByVal ElemType As Byte, ByVal Size As Long)
  Debug.Assert Not m_init
  m_init = True
  m_elemType = ElemType
  m_size = Size
End Sub

Public Property Get ElemType() As Byte
  Debug.Assert m_init
  ElemType = m_elemType
End Property

Public Property Get Size() As Long
  Debug.Assert m_init
  Size = m_size
End Property
