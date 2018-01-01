Attribute VB_Name = "FontHelper"
Option Explicit

Private Const DEFAULT_TEXT_SIZE As Long = 17

Private m_font                  As GDIPFont

Private m_defaultFontBold       As GDIPFont

Private m_blackBrush            As GDIPBrush

Private m_whiteBrush            As GDIPBrush
'Private m_fontFamily As GDIPFontFamily
'Private m_graphics As GDIPGraphics
'Private m_blankBitmap As GDIPBitmap
'Private m_colorMaker As GDIPColor

Public Function Dispose()

    If m_font Is Nothing Then

        Exit Function

    End If
    
    m_font.Dispose
    m_blackBrush.Dispose
End Function

Public Function GetWhiteBrush() As GDIPBrush
    InitializeIfRequired
    Set GetWhiteBrush = m_whiteBrush
End Function

Public Function GetBlackBrush() As GDIPBrush
    InitializeIfRequired
    Set GetBlackBrush = m_blackBrush
End Function

Public Function AppDefaultFont(Optional theStyle As FontStyle = FontStyleRegular) As GDIPFont
    InitializeIfRequired
    
    If theStyle = FontStyleRegular Then
        Set AppDefaultFont = m_font
    ElseIf theStyle = FontStyleBold Then
        Set AppDefaultFont = m_defaultFontBold
    Else
        LogError 0, "AppDefaultFont", "FontHelper", "Developer Error: Requested style not yet implemented"
    End If

End Function

Public Function AppFontFamily() As GDIPFontFamily
    InitializeIfRequired
    'Set AppFontFamily = m_fontFamily
End Function

Private Function InitializeIfRequired()

    If Not m_font Is Nothing Then

        Exit Function

    End If
    
    'Set m_colorMaker = New Colour
    Set m_font = New GDIPFont
    Set m_defaultFontBold = New GDIPFont
    
    'Set m_fontFamily = New GDIPFontFamily
    Set m_blackBrush = New GDIPBrush
    
    Set m_whiteBrush = New GDIPBrush
    'Set m_graphics = New GDIPGraphics
    'Set m_blankBitmap = New GDIPBitmap

    m_blackBrush.Colour.SetColourByHex ThemeHelper.Colour1.WebColor
    m_whiteBrush.Colour.SetColourByHex ThemeHelper.Colour2.WebColor
    
    Set m_font = ThemeHelper.Font
    Set m_defaultFontBold = ThemeHelper.FontBold
    
    'm_font.Constructor ThemeHelper.DefaultFace, ThemeHelper.FontSize, FontStyleRegular
    'm_defaultFontBold.Constructor ThemeHelper.DefaultFace, ThemeHelper.FontSize, FontStyleBold
    
    'm_fontFamily.Constructor ThemeHelper.DefaultFace
End Function

