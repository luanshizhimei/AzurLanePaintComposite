VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGDIpGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'// 声明API //
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, Bitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus" (ByVal Image As Long, Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus" (ByVal Image As Long) As Long
Private Declare Function GdipGraphicsClear Lib "GdiPlus" (ByVal Graphics As Long, ByVal lColor As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus" (ByVal Graphics As Long, ByVal Image As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Public Enum GpUnit
   UnitWorld
   UnitDisplay
   UnitPixel
   UnitPoint
   UnitInch
   UnitDocument
   UnitMillimeter
End Enum
'// 类变量 //
Private m_Bitmap As Long
Private m_Graphics As Long
Private m_Width As Long
Private m_Height As Long

Public Property Get Pointer() As Long
  Pointer = m_Graphics
End Property

Public Property Get Bitmap() As Long
  Bitmap = m_Bitmap
End Property

Public Function Create(ByVal l_Width As Long, ByVal L_Height As Long) As Object
  m_Width = l_Width: m_Height = L_Height
  GdipCreateBitmapFromScan0 m_Width, m_Height, 0, &H26200A, ByVal 0, m_Bitmap 'PixelFormat32bppARGB = &H26200A
  GdipGetImageGraphicsContext m_Bitmap, m_Graphics '以新建的bitmap作为绘图场景
  Set Create = Me
End Function

Public Sub Dispose()
  GdipDeleteGraphics m_Graphics
  GdipDisposeImage m_Bitmap
End Sub

Public Sub Clear()
  GdipGraphicsClear m_Bitmap, 0
End Sub

Public Sub DrawImage(ByVal Image As Object, _
  ByVal SrcX As Double, ByVal SrcY As Double, ByVal SrcW As Double, ByVal SrcH As Double, _
  ByVal DstX As Double, ByVal DstY As Double, ByVal DstW As Double, ByVal DstH As Double _
)
  GdipDrawImageRectRectI _
    m_Graphics, Image.Pointer, _
    DstX, DstY, DstW, DstH, _
    SrcX, SrcY, SrcW, SrcH, _
    UnitPixel
End Sub
