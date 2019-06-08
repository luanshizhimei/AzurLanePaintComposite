VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Width As Long
Private m_Height As Long
Private m_Depth As Long
Private m_V_Array() As Long
Private m_V_Space As Integer
Private m_V_Length As Integer
Private m_Vt_Array() As Double
Private m_Vt_Space As Integer
Private m_Vt_Length As Integer
Private m_Face_Array() As Object
Private m_Face_Space As Integer
Private m_Face_Length As Integer

Public Property Get Width() As Long
  Width = m_Width
End Property

Public Property Get Height() As Long
  Height = m_Height
End Property

Public Property Get Depth() As Long
  Depth = m_Depth
End Property

Public Property Get V(ByVal Num As Integer, ByVal ltype As Integer) As Long
  If (Num <= m_V_Length) Then
    V = m_V_Array(ltype, Num)
  End If
End Property

Public Property Get Vt(ByVal Num As Integer, ByVal ltype As Integer) As Double
  If (Num <= m_V_Length) Then
    Vt = m_Vt_Array(ltype, Num)
  End If
End Property

Public Property Get FaceCout() As Integer
  FaceCout = m_Face_Length
End Property

Public Property Get Face(ByVal Num As Integer) As Object
  If Num <= m_Face_Length Then Set Face = m_Face_Array(Num)
End Property

Public Function Load(ByVal FileName As String) As Boolean
  On Error GoTo 0
  If Len(Dir(FileName)) Then
    Dim nHandle As Long
    nHandle = FreeFile
    Open FileName For Input As #nHandle
    Dim Temp As String, Str() As String, Text() As String
    Dim D1 As Long, D2 As Long, D3 As Long, I As Integer, T As Integer
    Do Until EOF(nHandle)
      Line Input #nHandle, Temp
      Str = Split(Temp, " ")
      Select Case LCase(Str(0))
        Case "v"
          D1 = CLng(Str(1)): D2 = CLng(Str(2)): D3 = CLng(Str(3))
          If Abs(D1) > Abs(m_Width) Then m_Width = D1
          If Abs(D2) > Abs(m_Height) Then m_Height = D2
          If Abs(D3) > Abs(m_Depth) Then m_Depth = D3
  
          m_V_Length = m_V_Length + 1
          If m_V_Length > m_V_Space Then
            m_V_Space = m_V_Space + 100
            ReDim Preserve m_V_Array(2, m_V_Space)
          End If
          m_V_Array(0, m_V_Length) = D1
          m_V_Array(1, m_V_Length) = D2
          m_V_Array(2, m_V_Length) = D3
        Case "vt"
          m_Vt_Length = m_Vt_Length + 1
          If m_Vt_Length > m_Vt_Space Then
            m_Vt_Space = m_Vt_Space + 100
            ReDim Preserve m_Vt_Array(1, m_Vt_Space)
          End If
          m_Vt_Array(0, m_Vt_Length) = CDbl(Str(1))
          m_Vt_Array(1, m_Vt_Length) = CDbl(Str(2))
        Case "f"
          m_Face_Length = m_Face_Length + 1
          If m_Face_Length > m_Face_Space Then
            m_Face_Space = m_Face_Space + 100
            ReDim Preserve m_Face_Array(m_Face_Space)
          End If
          Set m_Face_Array(m_Face_Length) = New CObjFileFace
          For I = 1 To UBound(Str)
            Text = Split(Str(I), "/")
            m_Face_Array(m_Face_Length).Add Text
          Next
      End Select
    Loop
    Close nHandle
    Load = True
  End If
End Function

Private Sub Class_Terminate()
  Erase m_V_Array
  Erase m_Vt_Array
  Erase m_Face_Array
End Sub
