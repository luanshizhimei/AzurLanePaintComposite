VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGDIpImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'// 声明API //
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus" (ByVal FileName As Long, Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus" (ByVal Image As Long) As Long

'// 类变量 //
Private m_Width As Long
Private m_Height As Long
Private m_pImg As Long

Public Property Get Pointer() As Long
  Pointer = m_pImg
End Property

Public Property Get Width() As Long
  Width = m_Width
End Property

Public Property Get Height() As Long
  Height = m_Height
End Property

Public Function Load(ByRef l_Path As String) As Object
  GdipLoadImageFromFile StrPtr(l_Path), m_pImg
  GdipGetImageWidth m_pImg, m_Width
  GdipGetImageHeight m_pImg, m_Height
  Set Load = Me
End Function

Public Sub Dispose()
  GdipDisposeImage m_pImg
End Sub
