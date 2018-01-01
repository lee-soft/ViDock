VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayerdWindowHandles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public theDC      As Long

Public oldBitmap  As Long

Public mainBitmap As Long

Private winSize   As Win.SIZEL

Private srcPoint  As Win.POINTL

Private released  As Boolean

Private m_myBlend As BLENDFUNCTION

Public Property Get DC() As Long
    DC = theDC
End Property

Public Function Update(ByVal hWnd As Long, _
                       ByVal hdc As Long, _
                       Optional ByVal srcHdc As Long)

    If srcHdc = 0 Then srcHdc = theDC
    
    Call UpdateLayeredWindow(hWnd, hdc, ByVal 0&, winSize, srcHdc, srcPoint, 0, m_myBlend, ULW_ALPHA)
End Function

Public Function SetSize(newSize As Win.SIZEL)
    winSize = newSize
End Function

Public Function GetSize() As Win.SIZEL
    GetSize = winSize
End Function

Public Function SetPoint(newPoint As Win.POINTL)
    srcPoint = newPoint
End Function

Public Function GetPoint() As Win.POINTL
    GetPoint = srcPoint
End Function

Private Sub Class_Initialize()

    With m_myBlend
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    released = False
End Sub

Public Function SelectVBBitmap()
    'SelectObject theDC, oldBitmap
End Function

Public Sub Release()

    If released Then Exit Sub
    released = True

    SelectObject theDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC theDC
End Sub

Private Sub Class_Terminate()
    Release
End Sub
