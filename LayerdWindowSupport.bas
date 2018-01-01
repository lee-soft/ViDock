Attribute VB_Name = "LayerdWindowSupport"
Option Explicit

Private Const MODULE_NAME = "LayeredWindowSupport" 'TODO put these in all modules
                                                   'For conveniant logging

Public Declare Function UpdateLayeredWindow _
               Lib "user32.dll" (ByVal hWnd As Long, _
                                 ByVal hdcDst As Long, _
                                 pptDst As Any, _
                                 psize As Any, _
                                 ByVal hdcSrc As Long, _
                                 pptSrc As Any, _
                                 ByVal crKey As Long, _
                                 ByRef pblend As BLENDFUNCTION, _
                                 ByVal dwFlags As Long) As Long

Private m_layeredAttrBank As Collection

Public Const ULW_ALPHA = &H2

Public Const WS_EX_LAYERED = &H80000

Public Const AC_SRC_ALPHA As Long = &H1

Public Const AC_SRC_OVER = &H0

Public Type BLENDFUNCTION

    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte

End Type

Public Function IsLayeredWindow(ByVal hWnd As Long) As Boolean

    Dim WinInfo As Long

    WinInfo = GetWindowLong(hWnd, GWL_EXSTYLE)

    If (WinInfo And WS_EX_LAYERED) = WS_EX_LAYERED Then
        IsLayeredWindow = True
    Else
        IsLayeredWindow = False
    End If

End Function

Public Function UnMakeLayeredWindow(ByRef sourceForm As Form, _
                                    ByRef layeredData As LayerdWindowHandles)
    
    Dim winStyle As Long

    winStyle = GetWindowLong(sourceForm.hWnd, GWL_EXSTYLE)
    winStyle = winStyle And Not WS_EX_LAYERED
    
    SetWindowLong sourceForm.hWnd, GWL_EXSTYLE, winStyle
    
    sourceForm.Refresh
    
    'layeredData.ManualRelease
End Function

Public Function MakeLayerdWindow(ByRef sourceForm As Form, _
                                 Optional fromExistingLayeredWindow As Boolean = True, _
                                 Optional clickThrough As Boolean = False) As LayerdWindowHandles

    Dim KeyName As String

    KeyName = sourceForm.hWnd & "_hwnd"

    If m_layeredAttrBank Is Nothing Then
        Set m_layeredAttrBank = New Collection
    End If
    
    If ExistInCol(m_layeredAttrBank, KeyName) Then
        If fromExistingLayeredWindow Then
            m_layeredAttrBank(KeyName).Release
            m_layeredAttrBank.Remove KeyName
        Else
            Set MakeLayerdWindow = m_layeredAttrBank(KeyName)
            Call SetWindowLong(sourceForm.hWnd, GWL_EXSTYLE, GetWindowLong(sourceForm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)

            Exit Function

        End If
    End If

    Dim srcPoint   As Win.POINTL

    Dim winSize    As Win.SIZEL

    Dim mDC        As Long

    Dim tempBI     As BITMAPINFO

    Dim mainBitmap As Long

    Dim oldBitmap  As Long

    Dim theHandles As New LayerdWindowHandles

    Dim newStyle   As Long

    m_layeredAttrBank.Add theHandles, sourceForm.hWnd & "_hwnd"

    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32    ' Each pixel is 32 bit's wide
        .biHeight = sourceForm.ScaleHeight  ' Height of the form
        .biWidth = sourceForm.ScaleWidth    ' Width of the form
        .biPlanes = 1   ' Always set to 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
    End With
    
    mDC = CreateCompatibleDC(sourceForm.hdc)
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    
    If mainBitmap = 0 Then
        LogError 0, "MakeLayerdWindow", MODULE_NAME, "CreateDIBSection Failed"

        Exit Function

    End If
    
    oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected
    
    If oldBitmap = 0 Then

        'MsgBox "SelectObject Failed", vbCritical
        Exit Function

    End If
    
    newStyle = GetWindowLong(sourceForm.hWnd, GWL_EXSTYLE)
    newStyle = newStyle Or WS_EX_LAYERED
    
    If (clickThrough) Then
        newStyle = newStyle Or WS_EX_TRANSPARENT
    End If
    
    If SetWindowLong(sourceForm.hWnd, GWL_EXSTYLE, newStyle) = 0 Then
        'MsgBox "Failed to create layered window!"
        'Exit Function
    End If
    
    ' Needed for updateLayeredWindow call
    srcPoint.X = 0
    srcPoint.Y = 0
    winSize.cx = sourceForm.ScaleWidth
    winSize.cy = sourceForm.ScaleHeight
    
    theHandles.mainBitmap = mainBitmap
    theHandles.oldBitmap = oldBitmap
    theHandles.theDC = mDC
    
    theHandles.SetSize winSize
    theHandles.SetPoint srcPoint
    'theHandles.
    Set MakeLayerdWindow = theHandles

Handler:
End Function

