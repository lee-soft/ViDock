VERSION 5.00
Begin VB.Form ListMenu 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ListMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'! design with optional visible selector in the center (like with popup)

Private m_buttonMargin            As MARGIN

Private m_MarginText              As MARGIN

Private m_seperatorMargin         As MARGIN

Private m_shadowMargin            As MARGIN

Private m_tempGraphics            As GDIPGraphics

Private m_graphics                As GDIPGraphics

Private m_backgroundImage         As GDIPImage

Private m_buttonImage             As GDIPImage

Private m_listMenu                As Collection

Private m_layeredWindowProperties As LayerdWindowHandles

Private m_slices                  As Collection

Private m_buttonSlices            As Collection

Private m_seperator               As GDIPImage

Private m_centerX                 As Long

Private m_pointerSlice            As Slice

Private m_pointerY                As Long

Private m_pointerPosition         As Long

Private m_selectedItem            As Object

Private m_hoveredItem             As Object

Private m_selectedItemPosition    As Long

Private m_activated               As Boolean

Private m_showPointer             As Boolean

Private m_mouseTracking           As Boolean

Public ChildMenu                  As ListMenu

Public ParentMenu                 As ListMenu

Public Event onClosed()

Implements IHookSink

Sub ShowList(ByRef theList As Collection, _
             Y As Long, _
             X As Long, _
             Optional alignmentProfile As MenuControlConstants = vbPopupMenuCenterAlign, _
             Optional showPointer As Boolean = True, _
             Optional showFullBorder As Boolean = False)

    If showFullBorder Then
        Set m_MarginText = ThemeHelper.ListWindowTextMargin
        Set m_slices = ThemeHelper.ListWindow
        Set m_buttonImage = ThemeHelper.ListWindowImage
    Else
        Set m_MarginText = ThemeHelper.ListWindowTextClippedMargin
        Set m_slices = ThemeHelper.ListWindowClipped
        Set m_buttonImage = ThemeHelper.ListWindowClippedImage
    End If
    
    If Not m_listMenu Is Nothing Then
        If m_listMenu Is theList Then

            Exit Sub

        End If
    End If
    
    If Not ChildMenu Is Nothing Then
        Debug.Print "ShowList closeMe"
        ChildMenu.closeMe
    End If
        
    X = X - m_shadowMargin.Width
    
    'Me.Hide
    
    Set m_listMenu = theList
    
    Me.Width = (FindMaxWidth(theList, m_tempGraphics, FontHelper.AppDefaultFont) + (50)) * Screen.TwipsPerPixelX
    Me.Height = (FindMaxHeight(theList, MenuBarHelper.TEXTMODE_ITEM_Y_GAP) + (50)) * Screen.TwipsPerPixelY
    
    m_showPointer = showPointer
    
    If showPointer Then
        m_pointerPosition = (Me.ScaleWidth / 2) - (m_pointerSlice.Image.Width / 2)
    End If

    Me.Left = (X * Screen.TwipsPerPixelX)
    
    If alignmentProfile = vbPopupMenuCenterAlign Then
        m_centerX = (Me.ScaleWidth / 2)
        Me.Left = Me.Left - (m_centerX * Screen.TwipsPerPixelX)

        If (Me.Left + Me.Width) > Screen.Width Then
            Me.Left = Screen.Width - Me.Width
        End If
    End If
    
    Me.Top = Y
    
    InitializeGraphics
    
    Repaint
    'ShowWindow Me.hWnd, SW_SHOW
    
    If Not Me.Visible Then Me.Show
    'Me.Show
    
    m_activated = True
End Sub

Private Sub Form_Activate()
    Debug.Print "ListMenu_Activate!"
    
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Public Property Set ListMenu(ByRef newList As Collection)
    Set m_listMenu = newList
End Property

Private Sub Form_Deactivate()

    If Not m_activated Then Exit Sub

    If ChildMenu Is Nothing Then
        If Not ParentMenu Is Nothing Then
            ParentMenu.closeAncestors
        End If
        
        Debug.Print "Form_Deactivate closeMe"
        closeMe
    End If
    
    Exit Sub
    
    If m_activated Then
    
        If Not ParentMenu Is Nothing Then
            'ParentMenu.closeMe
        End If
    End If

End Sub

Private Sub Form_Initialize()
    Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    Set m_backgroundImage = New GDIPImage
    Set m_buttonImage = New GDIPImage
    
    Set m_tempGraphics = New GDIPGraphics
    
    Set m_buttonMargin = ThemeHelper.GetMargin("list_window_button")
    Set m_MarginText = ThemeHelper.ListWindowTextMargin
    Set m_seperatorMargin = ThemeHelper.GetMargin("list_window_seperator")
    Set m_shadowMargin = ThemeHelper.GetMargin("list_window_shadow")
    
    Set m_slices = ThemeHelper.ListWindowClipped
    Set m_buttonImage = ThemeHelper.ListWindowClippedImage
    
    Set m_buttonSlices = MenuListHelper.CreateButtonFromXML("list_window_button", m_buttonImage)
    Set m_seperator = ThemeHelper.Seperator
    
    m_tempGraphics.FromImage m_buttonImage
    
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
    'm_graphics.FromHDC Me.hdc
    
    'm_graphics.TextRenderingHint = TextRenderingHintAntiAlias
    m_graphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    m_graphics.SmoothingMode = SmoothingModeHighQuality
    m_graphics.PixelOffsetMode = PixelOffsetModeHighQuality
    'm_graphics.CompositingMode = CompositingModeSourceCopy
    
    'm_graphics.CompositingMode = CompositingModeSourceCopy
    'm_graphics.CompositingQuality = CompositingQualityHighQuality
    m_graphics.InterpolationMode = InterpolationModeNearestNeighbor
    
    'm_graphics.SetGDIFontRendering
End Function

Sub DrawListText()

    'On Error GoTo Handler
    If m_listMenu Is Nothing Then

        Exit Sub

    End If

    Dim thisItem     As Object

    Dim Y            As Long

    Dim X            As Long

    Dim nullSelected As Boolean

    Dim nullHovered  As Boolean

    Dim drawType     As Long

    Y = m_MarginText.Y
    X = m_MarginText.X

    For Each thisItem In m_listMenu

        If TypeOf thisItem Is Window Then
            thisItem.UpdateWindowText
        End If
        
        drawType = 0
        nullSelected = m_selectedItem Is Nothing
        nullHovered = m_hoveredItem Is Nothing
        
        If (Not nullSelected) And thisItem Is m_selectedItem Then
            drawType = ButtonPressed
        ElseIf (Not nullHovered) And thisItem Is m_hoveredItem Then
            drawType = ButtonUnpressed
        ElseIf TypeOf thisItem Is Window Then

            If thisItem.Flashing Then drawType = ButtonNotice
        End If
        
        If thisItem.Caption = "" Then
            m_graphics.DrawImage m_seperator, m_seperatorMargin.X, Y + m_seperatorMargin.Y, Me.ScaleWidth - m_seperatorMargin.X_Overflow, m_seperator.Height
        End If
        
        If drawType > 0 Then
            MenuListHelper.DrawButton m_buttonSlices, drawType, m_graphics, CreateRectL(m_buttonMargin.Height, Me.ScaleWidth - m_buttonMargin.X_Overflow, m_buttonMargin.X, Y - m_buttonMargin.Y_Overflow)
            m_graphics.DrawString thisItem.Caption, FontHelper.AppDefaultFont, GetWhiteBrush, CreatePointF(X, Y)
        Else
            m_graphics.DrawString thisItem.Caption, FontHelper.AppDefaultFont, GetBlackBrush, CreatePointF(X, Y)
        End If

        'm_Path.AddString thisWindow.Caption, m_fontFamily, 0, 12, CreateRectF(X, y + 3, 12, BorderWidth), 0
        'm_graphics.DrawString thisItem.Caption, FontHelper.AppDefaultFont, Brushes_Black, CreatePointF(X, y)

        Y = Y + MenuBarHelper.TEXTMODE_ITEM_Y_GAP
    Next

    'UpdateBuffer
    Exit Sub

Handler:
    Debug.Print "DrawList()" & Err.Description
End Sub

Sub Repaint()
    m_graphics.Clear

    Dim pointerArea As gdiplus.RECTF

    If Not m_pointerSlice Is Nothing And m_showPointer Then
        pointerArea = CreateRectF(CSng(m_pointerPosition), CSng(m_pointerY), CSng(m_pointerSlice.Image.Height), CSng(m_pointerSlice.Image.Width))
        m_graphics.Exclude pointerArea
    End If
    
    SliceHelper.DrawSlices m_slices, m_graphics, Me
    
    If Not m_pointerSlice Is Nothing Then
        'm_graphics.DrawRectangle SolidBlackPen, 0, 0, 30, 30
        m_graphics.ResetExclusions
        'm_graphics.DrawImage m_pointerSlice.Image, m_centerX - (m_pointerSlice.Image.Width / 2), Me.ScaleHeight - m_pointerSlice.Image.Height - 1, m_pointerSlice.Image.Width, m_pointerSlice.Image.Height
        m_graphics.DrawImageRectF m_pointerSlice.Image, pointerArea
    End If
    
    DrawListText
    
    'Me.Refresh
    m_layeredWindowProperties.Update Me.hWnd, m_layeredWindowProperties.theDC
End Sub

Private Sub Form_Load()
    HookWindow Me.hWnd, Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_hoveredItem Is Nothing Then

        Exit Sub

    End If
    
    Set m_selectedItem = m_hoveredItem
    Set m_hoveredItem = Nothing
    
    Debug.Print "Ere!"
    Repaint

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_mouseTracking = False Then
        m_mouseTracking = TrackMouseEvents(Me.hWnd)
    End If
    
    Dim selectedWindow As Object

    Set selectedWindow = GetSelectedItem(Y, m_selectedItemPosition)
    
    If selectedWindow Is Nothing Then Exit Sub
    If selectedWindow.Caption = "" Then
        Set m_hoveredItem = Nothing
        DeselectItem

        Exit Sub

    End If
    
    If m_hoveredItem Is selectedWindow Then

        Exit Sub

    End If
    
    Set m_hoveredItem = selectedWindow

    If m_hoveredItem Is Nothing Then

        Exit Sub

    End If
    
    Repaint

    If TypeOf m_hoveredItem Is MenuItem Then
        HandleMenuItemHovered m_hoveredItem, Me, m_selectedItemPosition, False
    End If

End Sub

Private Function GetSelectedItem(ByVal Y As Single, _
                                 Optional ByRef itemPosition As Long) As Object

    On Error GoTo Handler:
    
    Dim thisItem As Object

    Dim startY   As Long

    Dim endY     As Long

    startY = m_MarginText.Y
    itemPosition = 0

    For Each thisItem In m_listMenu

        endY = startY + MenuBarHelper.TEXTMODE_ITEM_Y_GAP
        itemPosition = itemPosition + 1

        If Y > startY And Y < endY Then
            Set GetSelectedItem = thisItem
            
            Exit For

        End If
        
        startY = endY
    Next
    
    Exit Function

Handler:
    LogError 0, "GetSelectedItem", "ListMenu", Err.Description
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim closeMenu As Boolean

    closeMenu = True

    If m_selectedItem Is Nothing Then Exit Sub
    
    If TypeOf m_selectedItem Is Window Then
    
        m_selectedItem.Flashing = False
        HandleWindow m_selectedItem.hWnd

    ElseIf TypeOf m_selectedItem Is ListFile Then
        
        Dim thisListFile As ListFile

        Set thisListFile = m_selectedItem
        
        HandleListFile thisListFile

    ElseIf TypeOf m_selectedItem Is MenuItem Then
        
        HandleMenuItem m_selectedItem, Me, m_selectedItemPosition, closeMenu
        
    End If
    
    Form_Deactivate
End Sub

Sub closeAncestors()
    m_activated = False
    Me.Hide
    
    Set m_listMenu = Nothing
    Set m_selectedItem = Nothing
    Set m_hoveredItem = Nothing
    
    If Not ChildMenu Is Nothing Then
        Unload ChildMenu
        Set ChildMenu = Nothing
    End If
    
    If Not ParentMenu Is Nothing Then
        ParentMenu.closeAncestors
        Set ParentMenu = Nothing
    End If
    
    RaiseEvent onClosed
End Sub

Sub closeMe()
    m_activated = False
    Me.Hide
    
    Set m_listMenu = Nothing
    Set m_selectedItem = Nothing
    Set m_hoveredItem = Nothing
    'Set m_pointerSlice = Nothing
    
    If Not ChildMenu Is Nothing Then
        Unload ChildMenu
        Set ChildMenu = Nothing
    End If
    
    If Not ParentMenu Is Nothing Then
        Set ParentMenu.ChildMenu = Nothing
        Set ParentMenu = Nothing
    End If
    
    RaiseEvent onClosed
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, _
                                      msg As Long, _
                                      wp As Long, _
                                      lp As Long) As Long

    On Error GoTo Handler

    If msg = WM_ACTIVATE Then
        If wp = WA_INACTIVE Then
            
            'Debug.Print lp & ":" & Me.hWnd
            
            Debug.Print "Taking focus:: " & GetFormByhWnd(lp).Name

            'Deactivate lp

            'If Not ChildMenu Is Nothing Then
            'If Not lp = ChildMenu.hWnd Then
            'closeMe
            'End If
            'Else
            'closeMe
            'End If

        End If

    ElseIf msg = WM_ACTIVATEAPP Then

        If LOWORD(wp) = WA_INACTIVE Then
            Debug.Print "hWndBelongToUs:: " & hWndBelongToUs(lp)
            Debug.Print "WM_ACTIVATEAPP closeMe"
            closeMe
        End If

    ElseIf msg = WM_MOUSELEAVE Then
        m_mouseTracking = False
        
        HandleMouseLeave
    End If
    
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = InvokeWindowProc(hWnd, msg, wp, lp)
End Function

Private Sub DeselectItem()
    Set m_hoveredItem = Nothing
    Repaint
    
End Sub

Private Sub HandleMouseLeave()
    DeselectItem
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnhookWindow Me.hWnd

    If Not ChildMenu Is Nothing Then
        Unload ChildMenu
    End If

End Sub

