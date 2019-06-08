VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjFileFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Dot_Array() As Object
Private m_Dot_Space As Integer
Private m_Dot_Length As Integer

Public Property Get Dot(ByVal Num As Integer) As Object
  Set Dot = m_Dot_Array(Num)
End Property

Public Function Add(ByRef lArray() As String)
  m_Dot_Length = m_Dot_Length + 1
  If m_Dot_Length > m_Dot_Space Then
    m_Dot_Space = m_Dot_Space + 2
    ReDim Preserve m_Dot_Array(m_Dot_Space)
  End If
  Set m_Dot_Array(m_Dot_Length) = New CObjFileDot
  
  Dim I As Integer: For I = 0 To UBound(lArray)
    Select Case I
      Case 0
        m_Dot_Array(m_Dot_Length).V = CLng(lArray(I))
      Case 1
        m_Dot_Array(m_Dot_Length).Vt = CLng(lArray(I))
      Case 2
        
    End Select
  Next
End Function

Private Sub Class_Terminate()
  Erase m_Dot_Array
End Sub
