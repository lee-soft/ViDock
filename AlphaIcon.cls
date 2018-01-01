VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlphaIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_Icon          As GDIPBitmap

Private m_IconXP        As GDIPBitmap

Public Property Get Image()

    If m_IconXP Is Nothing Then Exit Property
    Set Image = m_IconXP.Image
End Property

Public Function CreateFromHICON(ByVal icoHandle As Long)

    Dim ii         As ICONINFO

    Dim bmData     As BitmapData

    Dim bmBounds   As gdiplus.RECTL

    Dim PixelColor As ARGB

    pngIndex = pngIndex + 1

    Set m_Icon = New GDIPBitmap
    Set m_IconXP = New GDIPBitmap

    GetIconInfo icoHandle, ii
    m_Icon.CreateFromHBITMAP ii.hbmColor, 0

    bmBounds.Width = m_Icon.Image.Width
    bmBounds.Height = m_Icon.Image.Height

    bmData = m_Icon.LockBits(bmBounds, ImageLockModeRead, m_Icon.Image.ImgPixelFormat)

    m_IconXP.CreateFromSizeFormatData bmData.Height, bmData.Width, bmData.Stride, PixelFormat.Format32bppArgb, bmData.Scan0Ptr

    'check if this icon was really a 32Argb one or not
    'Explanation: icons which are not 32bit end up with all the alpha values being 0
    Dim i          As Long

    Dim X          As Long

    Dim Y          As Long
    
    Dim foundValid As Boolean
    
    If Not (m_IconXP.Image Is Nothing) Then

        Do While (i < m_IconXP.Image.Width * m_IconXP.Image.Height)
            Long2ARGB m_IconXP.GetPixel(X, Y), PixelColor

            If Not PixelColor.a = 0 Then
                foundValid = True

                Exit Do

            End If

            X = X + 1

            If (X = m_IconXP.Image.Width) Then
                X = 0
                Y = Y + 1
            End If

            i = i + 1
        Loop

    End If
    
    m_Icon.UnlockBits bmData
    DeleteObject ii.hbmColor
    DeleteObject ii.hbmMask
    
    If Not foundValid Then
        m_Icon.Dispose
        m_IconXP.Dispose
        
        m_IconXP.CreateFromHICON icoHandle
    End If

End Function

