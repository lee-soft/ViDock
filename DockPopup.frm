VERSION 5.00
Begin VB.Form DockPopup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17430
   LinkTopic       =   "Form1"
   ScaleHeight     =   63
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1162
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "DockPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_graphics                As GDIPGraphics

Private m_layeredWindowProperties As LayerdWindowHandles

Private m_backgroundImage         As GDIPImage

Private m_Path                    As GDIPGraphicPath

Private m_textPosition            As gdiplus.POINTF

Private m_caption                 As String

Private m_centerX                 As Long

Private m_pointerPosition         As Long

Private m_slices                  As Collection

Private m_pointerSlice            As Slice

Private m_pointerY                As Long

Public Function ShowTextPopup(ByVal theText As String, X As Long)
    'Me.Hide
    
    Debug.Print "Showing Text Popup: " & theText

    Dim ms             As gdiplus.RECTF

    Dim BorderWidth    As Long: BorderWidth = FindBorderWidth(m_slices) * 2

    Dim borderHeight   As Long: borderHeight = FindBorderHeight(m_slices) * 2

    Dim proposedHeight As Long

    ms = m_graphics.MeasureString(theText, AppDefaultFont)
    Debug.Print "MeasureStringWidth:: " & m_graphics.MeasureStringWidth(theText, AppDefaultFont)

    Me.Width = (ms.Width + BorderWidth) * Screen.TwipsPerPixelX
    proposedHeight = (ms.Height + borderHeight - 15)
    
    If proposedHeight < borderHeight Then
        proposedHeight = borderHeight
    End If
    
    Me.Height = proposedHeight * Screen.TwipsPerPixelY
    
    Me.Left = (X * Screen.TwipsPerPixelX)
    m_centerX = (Me.ScaleWidth / 2)
    
    Me.Left = Me.Left - (m_centerX * Screen.TwipsPerPixelX)
    m_pointerPosition = m_centerX - (m_pointerSlice.Image.Width / 2)
    
    If Me.Left + Me.Width > Screen.Width Then
        m_pointerPosition = m_pointerPosition + (Me.Left - (Screen.Width - Me.Width)) / Screen.TwipsPerPixelX
        Me.Left = Screen.Width - Me.Width
    End If
    
    If (Me.ScaleWidth - m_pointerPosition) < 42 Then
        m_pointerPosition = Me.ScaleWidth - 42
    End If
    
    m_caption = theText
    
    InitializeGraphics
    Repaint
    'ShowWindow Me.hWnd, SW_SHOWNOACTIVATE
    Me.Show
End Function

Private Sub PreparePath()
    'm_path.AddString "Test", FontHelper.AppFontFamily, fontStyle.FontStyleRegular, 14, m_textPosition, StringFormatFlagsNoWrap
End Sub

Private Sub Form_Initialize()
    m_textPosition.X = 30
    m_textPosition.Y = 27
    m_caption = "{no text}"
    
    StayOnTop Me, True
    
    Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    Set m_backgroundImage = New GDIPImage
    Set m_Path = New GDIPGraphicPath
    
    Set m_slices = SliceHelper.CreateSlicesFromXML("tooltip", m_backgroundImage)

    If ExistInCol(m_slices, "pointer") Then
        Set m_pointerSlice = m_slices("pointer")
        m_pointerY = m_backgroundImage.Height - m_pointerSlice.Y
    End If
    
    InitializeGraphics
    Repaint
End Sub

Private Function InitializeGraphics()

    If Not m_graphics Is Nothing Then
        If Not m_layeredWindowProperties Is Nothing Then
            m_graphics.ReleaseHDC m_layeredWindowProperties.theDC
            m_layeredWindowProperties.Release
        End If
    End If

    Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    Set m_graphics = New GDIPGraphics
    m_graphics.FromHDC m_layeredWindowProperties.theDC
    m_graphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    m_graphics.SmoothingMode = SmoothingModeHighQuality
    m_graphics.PixelOffsetMode = PixelOffsetModeHighQuality
    'm_graphics.CompositingMode = CompositingModeSourceCopy
    'm_graphics.CompositingQuality = CompositingQualityHighQuality
    m_graphics.InterpolationMode = InterpolationModeNearestNeighbor
    'm_graphics.
End Function

Sub Repaint()
    m_graphics.Clear
    'm_graphics.DrawImage m_backgroundImage, 0, 0, m_backgroundImage.Width, m_backgroundImage.Height

    Dim pointerArea As gdiplus.RECTF

    If Not m_pointerSlice Is Nothing Then
        If Not m_pointerSlice.Image Is Nothing Then
            pointerArea = CreateRectF(CSng(m_pointerPosition), CSng(m_pointerY), CSng(m_pointerSlice.Image.Height), CSng(m_pointerSlice.Image.Width))
        End If
    End If
    
    m_graphics.Exclude pointerArea
    SliceHelper.DrawSlices m_slices, m_graphics, Me
    
    If Not m_pointerSlice Is Nothing Then
        'm_graphics.DrawRectangle SolidBlackPen, 0, 0, 30, 30
        m_graphics.ResetExclusions
        'm_graphics.DrawImage m_pointerSlice.Image, m_centerX - (m_pointerSlice.Image.Width / 2), Me.ScaleHeight - m_pointerSlice.Image.Height - 1, m_pointerSlice.Image.Width, m_pointerSlice.Image.Height
        m_graphics.DrawImageRectF m_pointerSlice.Image, pointerArea
    End If
    
    m_graphics.DrawString m_caption, AppDefaultFont, FontHelper.GetBlackBrush, m_textPosition
    
    m_layeredWindowProperties.Update Me.hWnd, m_layeredWindowProperties.theDC
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not m_graphics Is Nothing And Not m_layeredWindowProperties Is Nothing Then
        m_graphics.ReleaseHDC m_layeredWindowProperties.theDC
        m_layeredWindowProperties.Release
    End If
    
    'FontHelper.Dispose
    'GDIPlusDispose
End Sub
