VERSION 5.00
Begin VB.Form TrayPopup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00383838&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   154
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "TrayPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : TrayPopup
'    Project    : ViDock
'
'    Description: Popup when you mouse over the tray icon area
'
'--------------------------------------------------------------------------------
Option Explicit

Private WithEvents m_trayBar      As SystemTrayManager
Attribute m_trayBar.VB_VarHelpID = -1

Private m_squareRoot              As Long

Private m_graphics                As GDIPGraphics

Private m_layeredWindowProperties As LayerdWindowHandles

Private m_backgroundImage         As GDIPImage

Private m_Path                    As GDIPGraphicPath

Private m_centerX                 As Long

Private m_pointerPosition         As Long

Private m_slices                  As Collection

Private m_pointerSlice            As Slice

Private m_pointerY                As Long

Private m_marginY                 As Long

Private m_marginX                 As Long

Private m_mouseTracking           As Boolean

Implements IHookSink

Public Function ShowTrayPopup(ByVal X As Long)

    '    Me.Hide

    Dim ms             As gdiplus.RECTF

    Dim BorderWidth    As Long: BorderWidth = FindBorderWidth(m_slices) * 2

    Dim borderHeight   As Long: borderHeight = FindBorderHeight(m_slices) * 2

    Dim proposedHeight As Long

    m_marginY = borderHeight / 2
    m_marginX = BorderWidth / 2

    m_trayBar.Update
    m_squareRoot = MathHelper.Ceiling(MathHelper.Sqrt(m_trayBar.CountIcons))
    m_trayBar.ColumnLimit = m_squareRoot

    Me.Width = (BorderWidth + ((ICON_SIZE + MARGIN) * m_squareRoot)) * Screen.TwipsPerPixelX
    
    proposedHeight = (ms.Height + borderHeight - 15) + ((ICON_SIZE + MARGIN) * m_squareRoot)
    
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
    
    InitializeGraphics
    Repaint

    'ShowWindow Me.hWnd, SW_SHOWNORMAL
    Me.Show
            
    'Me.Left = 0
End Function

Private Sub Form_DblClick()
    'Dim MousePosition As points
    'MousePosition.X = X - m_marginX
    'MousePosition.Y = Y - m_marginY
    
    'If MouseInsideObject(MousePosition, m_trayBar.Dimensions.Left, m_trayBar.Dimensions.Top, m_trayBar.Dimensions.Width, m_trayBar.Dimensions.Height) Then
    m_trayBar.MouseDblClick
    'Exit Sub

End Sub

Private Sub Form_Initialize()
    StayOnTop Me, True
    
    Set m_trayBar = New SystemTrayManager

    Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    Set m_backgroundImage = New GDIPImage
    Set m_Path = New GDIPGraphicPath
    Set m_trayBar.HostForm = Me
    
    Set m_slices = SliceHelper.CreateSlicesFromXML("tray_popup", m_backgroundImage)

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

Private Sub Form_Load()
    'Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    HookWindow Me.hWnd, Me
End Sub

Sub Repaint()
    m_graphics.Clear
    'm_graphics.DrawImage m_backgroundImage, 0, 0, m_backgroundImage.Width, m_backgroundImage.Height

    Dim pointerArea As gdiplus.RECTF

    If Not m_pointerSlice Is Nothing Then
        pointerArea = CreateRectF(CSng(m_pointerPosition), CSng(m_pointerY), CSng(m_pointerSlice.Image.Height), CSng(m_pointerSlice.Image.Width))
    End If
    
    m_graphics.Exclude pointerArea
    SliceHelper.DrawSlices m_slices, m_graphics, Me
    
    If Not m_pointerSlice Is Nothing Then
        'm_graphics.DrawRectangle SolidBlackPen, 0, 0, 30, 30
        m_graphics.ResetExclusions
        'm_graphics.DrawImage m_pointerSlice.Image, m_centerX - (m_pointerSlice.Image.Width / 2), Me.ScaleHeight - m_pointerSlice.Image.Height - 1, m_pointerSlice.Image.Width, m_pointerSlice.Image.Height
        m_graphics.DrawImageRectF m_pointerSlice.Image, pointerArea
    End If
    
    'm_graphics.DrawImage m_trayBar.Image, m_marginX, m_marginY, m_trayBar.Image.Width, m_trayBar.Image.Height
    m_graphics.DrawImage m_trayBar.Image, CSng(m_marginX), CSng(m_marginY), m_trayBar.Image.Width, m_trayBar.Image.Height
    
    m_layeredWindowProperties.Update Me.hWnd, m_layeredWindowProperties.theDC
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim MousePosition As POINTS

    MousePosition.X = X - m_marginX
    MousePosition.Y = Y - m_marginY
    
    If MouseInsideObject(MousePosition, m_trayBar.Dimensions.Left, m_trayBar.Dimensions.Top, m_trayBar.Dimensions.Width, m_trayBar.Dimensions.Height) Then
        m_trayBar.MouseDown Button, CSng(MousePosition.X), CSng(MousePosition.Y)

        Exit Sub

    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_mouseTracking = False Then
        m_mouseTracking = TrackMouseEvents(Me.hWnd)
    End If

    Dim MousePosition As POINTS

    MousePosition.X = X - m_marginX
    MousePosition.Y = Y - m_marginY

    m_trayBar.MouseMove MousePosition
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim MousePosition As POINTS

    MousePosition.X = X - m_marginX
    MousePosition.Y = Y - m_marginY
    
    If MouseInsideObject(MousePosition, m_trayBar.Dimensions.Left, m_trayBar.Dimensions.Top, m_trayBar.Dimensions.Width, m_trayBar.Dimensions.Height) Then
        m_trayBar.MouseUp Button, CSng(MousePosition.X), CSng(MousePosition.Y)

        Exit Sub

    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    UnhookWindow Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not m_graphics Is Nothing And Not m_layeredWindowProperties Is Nothing Then
        m_graphics.ReleaseHDC m_layeredWindowProperties.theDC
        m_layeredWindowProperties.Release
    End If
    
    'FontHelper.Dispose
    'GDIPlusDispose
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, _
                                      msg As Long, _
                                      wp As Long, _
                                      lp As Long) As Long

    On Error GoTo Handler

    If msg = WM_ACTIVATEAPP Then
        If LOWORD(wp) = WA_INACTIVE Then
            If Not m_mouseTracking Then
                Me.Hide
            End If
        End If

    ElseIf msg = WM_MOUSELEAVE Then
        m_mouseTracking = False
    End If

Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = InvokeWindowProc(hWnd, msg, wp, lp)
End Function


