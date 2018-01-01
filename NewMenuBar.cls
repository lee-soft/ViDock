VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewMenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : NewMenuBar
'    Project    : ViDock
'
'    Description: The Menu Bar object. Manages the App Menu items
'
'--------------------------------------------------------------------------------
Option Explicit

Private Const BUTTON_TEXT_GAP As Long = 10

Private m_textPath            As GDIPGraphicPath

Private m_items               As Collection

Private m_graphics            As GDIPGraphics

Private m_dimensions          As gdiplus.RECTL

Private m_menuOver            As GDIPImage

Private m_menuBitmap          As GDIPBitmap

Private m_menuBackground      As GDIPImage

Private m_buttonSlices        As Collection

Private m_hoveredItem         As MenuItem

Private m_windowTitle         As String

Private m_rootItem            As MenuItem

Private m_targethWnd          As Long

Private m_bgFinderMargin      As MARGIN

Public HostForm As Form

Private WithEvents m_listMenu As ListMenu
Attribute m_listMenu.VB_VarHelpID = -1

Public Event onChanged()

Public Function Update()
    
    If GetWindowTextByhWnd(m_targethWnd) <> m_windowTitle Then
        PopulateFromhWnd m_targethWnd
    End If
    
End Function

Public Function ResetSelector()
    Set m_hoveredItem = Nothing
    ReDraw
End Function

Public Property Let Background(newBackground As GDIPImage)
    Set m_menuBackground = newBackground
    ReDraw
    
    'RaiseEvent onChanged
End Property

Public Property Get Image() As GDIPImage
    Set Image = m_menuBitmap.Image
End Property

Public Property Get Dimensions() As gdiplus.RECTL
    Dimensions = m_dimensions
End Property

Public Property Let Dimensions(newDimensions As gdiplus.RECTL)
    m_dimensions = newDimensions
    
    m_menuBitmap.CreateFromSizeFormat newDimensions.Width, newDimensions.Height, GDIPlusWrapper.Format32bppArgb
    
    Set m_graphics = New GDIPGraphics
    m_graphics.FromImage m_menuBitmap.Image
    
    m_graphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
    m_graphics.SmoothingMode = SmoothingModeHighQuality
    m_graphics.PixelOffsetMode = PixelOffsetModeHighQuality
    
    'm_graphics.CompositingQuality = CompositingQualityHighQuality
    m_graphics.InterpolationMode = InterpolationModeNearestNeighbor
End Property

Public Function MouseDown(Button As Integer, X As Single, Y As Single)

    If HostForm Is Nothing Then
        LogError 0, "MouseDown", "NewMenuBar", "HostForm unitialized!"
        Exit Function

    End If
    
    If m_hoveredItem Is Nothing Then

        Exit Function

    End If
    
    'm_listMenu.closeMe
    m_listMenu.ShowList m_hoveredItem.Children, HostForm.Top + ((m_dimensions.Top + m_dimensions.Height - m_bgFinderMargin.Height) * Screen.TwipsPerPixelY), (HostForm.Left / Screen.TwipsPerPixelX) + m_dimensions.Left + m_hoveredItem.Left - BUTTON_TEXT_GAP, vbPopupMenuLeftAlign, False
    
End Function

Public Function MouseMove(Position As POINTS)
    
    Dim newItem As MenuItem
    
    Set newItem = GetSelectedItem(Position.X)
    
    If newItem Is Nothing Then

        Exit Function

    End If
    
    If newItem Is m_hoveredItem Then

        Exit Function

    End If
    
    If newItem Is m_rootItem Then

        Exit Function

    End If
    
    Set m_hoveredItem = newItem
    
    If Not m_listMenu Is Nothing Then
        If m_listMenu.Visible Then
            m_listMenu.ShowList m_hoveredItem.Children, HostForm.Top + ((m_dimensions.Top + m_dimensions.Height - m_bgFinderMargin.Height) * Screen.TwipsPerPixelY), (HostForm.Left / Screen.TwipsPerPixelX) + m_dimensions.Left + m_hoveredItem.Left - BUTTON_TEXT_GAP, vbPopupMenuLeftAlign, False
        End If
    End If
    
    ReDraw
    RaiseEvent onChanged

End Function

Public Sub MouseLeft()

    If m_listMenu.Visible Then
        m_listMenu.Hide
    End If
    
    Debug.Print "MenuBar::MouseLeft"
End Sub

Private Function GetSelectedItem(ByVal X As Long) As MenuItem

    Dim thisItem As MenuItem

    For Each thisItem In m_items

        If X > thisItem.Left And X < thisItem.Left + thisItem.Width Then
            Set GetSelectedItem = thisItem

            Exit For

        End If

    Next

End Function

Public Function PopulateFromhWnd(ByVal hWnd As Long)

    Set m_items = New Collection
    m_targethWnd = hWnd
    m_windowTitle = GetWindowTextByhWnd(m_targethWnd)
    
    m_rootItem.Caption = m_windowTitle
    m_items.Add m_rootItem

    PopulateMenuFromHandle m_items, GetMenu(hWnd), hWnd
    ReDraw
End Function

Public Function PopulateFromTest()

    If m_graphics Is Nothing Then
        Debug.Print "PopulateFromTest:: No graphics object!"

        Exit Function

    End If

    Dim thisItem As New MenuItem

    thisItem.Caption = "File"

    m_items.Add thisItem
    
    Set thisItem = New MenuItem
    m_items.Add thisItem
    thisItem.Caption = "Edit"
    
    Set thisItem = New MenuItem
    m_items.Add thisItem
    thisItem.Caption = "View"
    
    Set thisItem = New MenuItem
    m_items.Add thisItem
    thisItem.Caption = "Find"
    
    ReDraw
End Function

Private Function PositionItems()
    
    Dim thisItem  As MenuItem

    Dim xPosition As Long

    xPosition = MENU_MARGIN_X
    
    For Each thisItem In m_items

        If thisItem Is m_rootItem Then
            thisItem.Width = m_graphics.MeasureStringWidth(thisItem.Caption, FontHelper.AppDefaultFont(FontStyleBold))
        Else
            thisItem.Width = m_graphics.MeasureStringWidth(thisItem.Caption, FontHelper.AppDefaultFont)
        End If
        
        thisItem.Left = xPosition
        xPosition = xPosition + thisItem.Width + ITEM_MARGIN_X
    Next
    
End Function

Private Function ReDraw()
    PositionItems
    
    Dim thisItem As MenuItem

    Set m_textPath = New GDIPGraphicPath
    
    m_graphics.Clear
    'm_graphics.DrawImage m_menuBackground, 0, 0, m_dimensions.Width, m_dimensions.Height
    m_graphics.DrawImage m_menuBackground, 0, 0, CSng(m_dimensions.Width), CSng(m_dimensions.Height)
    
    For Each thisItem In m_items

        If thisItem Is m_rootItem Then
            m_graphics.DrawString thisItem.Caption, FontHelper.AppDefaultFont(FontStyleBold), GetBlackBrush, CreatePointF(thisItem.Left, MENU_MARGIN_Y)
            
        ElseIf thisItem Is m_hoveredItem Then
            MenuListHelper.DrawButton m_buttonSlices, ButtonOver, m_graphics, CreateRectL(m_dimensions.Height, thisItem.Width + (BUTTON_TEXT_GAP * 2), thisItem.Left - BUTTON_TEXT_GAP, 0)
            m_graphics.DrawString thisItem.Caption, FontHelper.AppDefaultFont, GetWhiteBrush, CreatePointF(thisItem.Left, MENU_MARGIN_Y)
            
        Else
            MenuListHelper.DrawButton m_buttonSlices, ButtonUnpressed, m_graphics, CreateRectL(m_dimensions.Height, thisItem.Width, thisItem.Left, 0)
            m_graphics.DrawString thisItem.Caption, FontHelper.AppDefaultFont, GetBlackBrush, CreatePointF(thisItem.Left, MENU_MARGIN_Y)

        End If
        
        'm_textPath.AddString thisItem.Caption, FontHelper.AppFontFamily, FontStyleRegular, FontHelper.AppDefaultFont.Size, CreateRectF(thisItem.Left, m_dimensions.Top, 20, m_dimensions.Width), 0
    Next
    
    RaiseEvent onChanged
End Function

Private Sub Class_Initialize()
    Set m_textPath = New GDIPGraphicPath
    Set m_items = New Collection
    Set m_menuBitmap = New GDIPBitmap
    Set m_graphics = New GDIPGraphics
    Set m_menuOver = New GDIPImage
    Set m_listMenu = New ListMenu
    Set m_rootItem = New MenuItem

    Set m_buttonSlices = MenuListHelper.CreateButtonFromXML("menubar_states", m_menuOver)
    Set m_bgFinderMargin = GetMargin("bgfinder_shadow")
End Sub

Private Sub Class_Terminate()
    Unload m_listMenu
End Sub

Private Sub m_listMenu_onClosed()
    Set m_hoveredItem = Nothing
    ReDraw
End Sub
