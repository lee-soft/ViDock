VERSION 5.00
Begin VB.Form StartOrb 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "StartOrb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : StartOrb
'    Project    : ViDock
'
'    Description: The start menu launcher form
'
'--------------------------------------------------------------------------------
Option Explicit

Implements IHookSink

Private m_theStartButton          As GDIPImage

Private m_graphics                As GDIPGraphics

Private m_layeredWindowProperties As LayerdWindowHandles

Private m_buttonSlices            As Collection

Private m_buttonState             As ButtonState

Private m_mouseTracking           As Boolean

Public Event onMove(ByRef newX As Long, ByRef newY As Long)

Private Sub Form_Click()
    Me.Move 40, 40
End Sub

Private Sub Form_Load()
    HookWindow Me.hWnd, Me
    TrackMouseEvents Me.hWnd

    Set m_theStartButton = New GDIPImage
    Set m_graphics = New GDIPGraphics

    Set m_buttonSlices = MenuListHelper.CreateButtonFromXML("start_button", m_theStartButton)
    
    Dim thisCollection As Collection

    Dim thisSlice      As Slice

    Set thisCollection = m_buttonSlices(1)
    Set thisSlice = thisCollection(1)
    
    Me.Width = m_theStartButton.Width * Screen.TwipsPerPixelX
    Me.Height = thisSlice.Height * Screen.TwipsPerPixelY

    'StayOnTop Me, True
    
    InitializeGraphics
    Paint
End Sub

Private Function InitializeGraphics()
    Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    m_graphics.FromHDC m_layeredWindowProperties.theDC
End Function

Sub Paint()
    m_graphics.Clear
    MenuListHelper.DrawButton m_buttonSlices, m_buttonState, m_graphics, CreateRectL(54, Me.ScaleWidth, 0, 0)
    
    m_layeredWindowProperties.Update Me.hWnd, m_layeredWindowProperties.theDC
    'Me.Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        HandleMouseClicked
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_mouseTracking = False Then
        m_mouseTracking = TrackMouseEvents(Me.hWnd)
        HandleMouseEnter
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_graphics.ReleaseHDC m_layeredWindowProperties.theDC
    m_layeredWindowProperties.Release
    
    UnhookWindow Me.hWnd
End Sub

'Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'If iMsg = WM_MOUSELEAVE Then
'm_mouseTracking = False
'HandleMouseLeave
'End If

'End Function

Private Function HandleMouseClicked()
    m_buttonState = ButtonPressed
    Paint
    
    ShowStartMenu
End Function

Private Function HandleMouseLeave()
    m_buttonState = ButtonUnpressed
    Paint
End Function

Private Function HandleMouseEnter()
    m_buttonState = ButtonOver
    Paint
End Function

Private Function IHookSink_WindowProc(hWnd As Long, _
                                      msg As Long, _
                                      wp As Long, _
                                      lp As Long) As Long

    On Error GoTo Handler

    If msg = WM_MOUSELEAVE Then
        m_mouseTracking = False
        
        HandleMouseLeave
        
    ElseIf msg = WM_WINDOWPOSCHANGING Then
        
        Dim thisWindowPosition As windowPos
        
        CopyMemory thisWindowPosition, ByVal lp, LenB(thisWindowPosition)
        RaiseEvent onMove(thisWindowPosition.X, thisWindowPosition.Y)
        CopyMemory ByVal lp, thisWindowPosition, LenB(thisWindowPosition)
        
    End If

Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = InvokeWindowProc(hWnd, msg, wp, lp)
End Function
