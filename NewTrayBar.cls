VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewTrayBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_graphics As GDIPGraphics
Private m_dimensions As GdiPlus.RECTL
Private m_trayBitmap As GDIPBitmap
'Private m_trayBackground As GDIPImage

Public Event onChanged()

'Public Property Let Background(newBackground As GDIPImage)
    'Set m_menuBackground = newBackground
'End Property

Public Property Get Image() As GDIPImage
    Set Image = m_trayBitmap.Image
End Property

Public Property Get Dimensions() As GdiPlus.RECTL
    Dimensions = m_dimensions
End Property

Public Property Let Dimensions(newDimensions As GdiPlus.RECTL)
    m_dimensions = newDimensions
    
    m_trayBitmap.CreateFromSize newDimensions.Width, newDimensions.Height
    m_graphics.FromImage m_trayBitmap.Image
    
    m_graphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
    'm_graphics.SmoothingMode = SmoothingModeHighQuality
    'm_graphics.InterpolationMode = InterpolationModeHighQualityBicubic
    'm_graphics.PixelOffsetMode = PixelOffsetModeHighQuality
    'm_graphics.CompositingQuality = CompositingQualityHighQuality

    m_graphics.PixelOffsetMode = PixelOffsetModeHalf
    m_graphics.CompositingQuality = CompositingQualityHighSpeed
    m_graphics.InterpolationMode = InterpolationModeNearestNeighbor
End Property

Private Sub GetTaskbarItems()

Dim sysPagerHandle As Long
Dim trayHandle As Long
Dim trayToolbarHandle As Long
Dim itemCount As Long

Dim theButton As TBBUTTON
Dim buttonIndex As Long

    trayHandle = ShellHelper.TrayHandler
    If trayHandle = 0 Then
        Exit Sub
    End If

    sysPagerHandle = FindWindowEx(trayHandle, 0, "SysPager", vbNullString)
    If sysPagerHandle = 0 Then
        Exit Sub
    End If
    
    trayToolbarHandle = FindWindowEx(sysPagerHandle, 0, "ToolbarWindow32", vbNullString)
    If trayToolbarHandle = 0 Then
        Exit Sub
    End If

    trayToolbarHandle = FindTrayToolbarWindow
    itemCount = SendMessage(ByVal trayToolbarHandle, ByVal TB_BUTTONCOUNT, 0, 0)

    For buttonIndex = 0 To itemCount
        If SendMessage(ByVal trayToolbarHandle, ByVal TB_GETBUTTON, ByVal buttonIndex, theButton) = APIFALSE Then
            Debug.Print "failed!"
        End If
        
        'Debug.Print theButton.
    Next

End Sub

Private Sub Class_Initialize()
    Set m_graphics = New GDIPGraphics
    Set m_trayBitmap = New GDIPBitmap
    
    GetTaskbarItems
    
    ExitApplication
End Sub

