VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemTrayManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : SystemTrayManager
'    Project    : ViDock
'
'    Description: (Not related to the OS System Tray) manages all functionality
'                 to do with the dock manager
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_isExplorer64bit    As Boolean

Private m_trayhWnd           As Long

Private m_traypID            As Long

Private m_buttons            As Collection

Private m_bitmap             As GDIPBitmap

Private m_graphics           As GDIPGraphics

Private m_popup              As DockPopup

Private m_dimensions         As gdiplus.RECTL

Private m_lastSelectedButton As TrayButtonInfo

Private m_columnLimit        As Long

Private m_finalRowcount      As Long

Private m_lastIconCount      As Long

Public HostForm              As Form

Public Event onChange()

Public Function IsPopupShowing() As Boolean

    On Error GoTo Handler:
    
    IsPopupShowing = m_popup.Visible

    Exit Function

Handler:
End Function

Public Property Let ColumnLimit(newLimit As Long)
    
    m_columnLimit = newLimit

    UpdateCalculations
    UpdateImage
    
End Property

Private Function UpdateCalculations()

    If m_columnLimit = 0 Then Exit Function

    m_lastIconCount = CountIcons
    m_finalRowcount = MathHelper.Ceiling(m_lastIconCount / m_columnLimit)
    
End Function

Public Property Get Dimensions() As gdiplus.RECTL
    Dimensions.Top = m_dimensions.Top
    Dimensions.Left = m_dimensions.Left
    
    Dimensions.Width = m_bitmap.Image.Width
    Dimensions.Height = m_bitmap.Image.Height
End Property

Public Property Let Dimensions(newDimensions As gdiplus.RECTL)
    m_dimensions.Left = newDimensions.Left
    m_dimensions.Top = newDimensions.Top
    
End Property

Public Property Get Image() As GDIPImage
    Set Image = m_bitmap.Image
End Property

Public Property Get Buttons() As Collection
    Set Buttons = m_buttons
End Property

Private Sub Class_Initialize()

    Dim hProcess As Long

    Dim ret      As Long

    Set m_bitmap = New GDIPBitmap
    m_bitmap.CreateFromSizeFormat 1, ICON_SIZE, GDIPlusWrapper.Format32bppArgb
    
    Set m_graphics = New GDIPGraphics
    m_graphics.FromImage m_bitmap.Image

    m_trayhWnd = SystemTrayHelper.FindTrayToolbarWindowHiddenIcons

    If m_trayhWnd = -1 Then Exit Sub
    
    ret = GetWindowThreadProcessId(m_trayhWnd, m_traypID)
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, m_traypID)

    If hProcess = 0 Then Exit Sub
    
    m_isExplorer64bit = IsProcess64bit(hProcess)
    CloseHandle hProcess
    
    Set m_buttons = New Collection
    Update
    
End Sub

Public Property Let Popup(newPopup As DockPopup)
    Set m_popup = newPopup
End Property

Public Function Update()

    Dim trayInfo          As TrayButtonInfo

    Dim iconIndex         As Long

    Dim changed           As Boolean

    Dim buttonsThisUpdate As Collection

    Set buttonsThisUpdate = New Collection
    
    For iconIndex = 0 To SysTrayIconCount(m_trayhWnd)
        Set trayInfo = GetButtonInfo(iconIndex)
        
        If trayInfo.Visible Then
            If Not ExistInCol(buttonsThisUpdate, trayInfo.App_Identifier) Then buttonsThisUpdate.Add trayInfo, trayInfo.App_Identifier
        End If

    Next

    For Each trayInfo In m_buttons

        If Not ExistInCol(buttonsThisUpdate, trayInfo.App_Identifier) Then
            changed = True

            Exit For

        End If

    Next
    
    If Not changed Then

        For Each trayInfo In buttonsThisUpdate

            If Not ExistInCol(m_buttons, trayInfo.App_Identifier) Then
                changed = True

                Exit For

            End If

        Next

    End If
    
    If Not changed Then

        For Each trayInfo In buttonsThisUpdate

            If trayInfo.hIcon <> m_buttons(trayInfo.App_Identifier).hIcon Then
                changed = True

                Exit For

            End If

        Next

    End If
    
    If changed Then
        Set m_buttons = buttonsThisUpdate
        
        UpdateCalculations
        UpdateImage
        
        Debug.Print "onChanged!"
        RaiseEvent onChange
    End If

End Function

Private Function UpdateImage()

    Dim trayInfo      As TrayButtonInfo

    Dim buttonIndex   As Long

    Dim thisAlphaIcon As AlphaIcon

    Dim rowIndex      As Long

    Dim columnIndex   As Long

    Set thisAlphaIcon = New AlphaIcon
    
    If m_bitmap.Image.Width <> m_columnLimit * (ICON_SIZE + MARGIN) Or m_bitmap.Image.Height <> m_finalRowcount * (ICON_SIZE + MARGIN) Then
        
        Set m_bitmap = New GDIPBitmap
        'm_bitmap.CreateFromSize m_buttons.count * (SystemTrayHelper.ICON_SIZE + SystemTrayHelper.MARGIN), SystemTrayHelper.ICON_SIZE
        m_bitmap.CreateFromSizeFormat m_columnLimit * (ICON_SIZE + MARGIN), m_finalRowcount * (ICON_SIZE + MARGIN), GDIPlusWrapper.Format32bppArgb
        
        m_graphics.FromImage m_bitmap.Image
    End If
    
    m_graphics.Clear

    For Each trayInfo In m_buttons

        If trayInfo.Visible Then
            thisAlphaIcon.CreateFromHICON trayInfo.hIcon

            If columnIndex = m_columnLimit Then
                rowIndex = rowIndex + 1
                columnIndex = 0
            End If
            
            'm_graphics.DrawImage thisAlphaIcon.Image, (buttonIndex * (SystemTrayHelper.ICON_SIZE + SystemTrayHelper.MARGIN)), 0, SystemTrayHelper.ICON_SIZE, SystemTrayHelper.ICON_SIZE
            m_graphics.DrawImage thisAlphaIcon.Image, columnIndex * (ICON_SIZE + MARGIN), rowIndex * (ICON_SIZE + MARGIN), ICON_SIZE, ICON_SIZE
            
            columnIndex = columnIndex + 1
            buttonIndex = buttonIndex + 1
        End If

    Next

End Function

Public Function CountIcons() As Long
    CountIcons = SysTrayIconCount(m_trayhWnd)
End Function

Private Function GetButtons() As Collection

    Dim trayInfo  As TrayButtonInfo

    Dim iconIndex As Long

    Dim Buttons   As Collection

    Set Buttons = New Collection
    Set GetButtons = Buttons
    
    Dim sysTrayIconCountSnapshot As Long
    sysTrayIconCountSnapshot = SysTrayIconCount(m_trayhWnd)

    For iconIndex = 0 To sysTrayIconCountSnapshot
    
        Set trayInfo = GetButtonInfo(iconIndex)

        If Not ExistInCol(Buttons, trayInfo.App_Identifier) Then Buttons.Add trayInfo, trayInfo.App_Identifier
    Next

End Function

Private Function GetButtonInfo(iIndex As Long) As TrayButtonInfo

    If m_isExplorer64bit Then
        Set GetButtonInfo = SysTrayGetButtonInfo_64(m_trayhWnd, m_traypID, iIndex)
    Else
        Set GetButtonInfo = SysTrayGetButtonInfo_86(m_trayhWnd, m_traypID, iIndex)
    End If

End Function

Public Function MouseDblClick()

    If m_lastSelectedButton Is Nothing Then

        Exit Function

    End If
    
    Dim cursorPos As POINTL

    GetCursorPos cursorPos

    PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal MAKEWPARAM(cursorPos.X, cursorPos.Y), ByVal WM_LBUTTONDBLCLK
End Function


Public Function MouseUp(Button As Integer, X As Single, Y As Single)
    
    Dim cursorPos As POINTL

    GetCursorPos cursorPos
    
    If m_lastSelectedButton Is Nothing Then

        Exit Function

    End If

    If Button = vbRightButton Then
        'WM_NCRBUTTONDOWN
        
        'PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal 0, ByVal WM_NCRBUTTONUP
        'PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal m_lastSelectedButton.uCallbackMessage, ByVal WM_RBUTTONUP
        PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal MAKEWPARAM(cursorPos.X, cursorPos.Y), ByVal WM_CONTEXTMENU
    ElseIf Button = vbLeftButton Then
        PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal MAKEWPARAM(cursorPos.X, cursorPos.Y), ByVal WM_LBUTTONUP
    End If
End Function

Public Function MouseDown(Button As Integer, X As Single, Y As Single)

    On Error GoTo Handler
    
    If m_lastSelectedButton Is Nothing Then

        Exit Function

    End If
    
    Dim mainhWnd  As Long

    Dim winStyle  As Long

    Dim cursorPos As POINTL

    GetCursorPos cursorPos
    
    'mainhWnd = GetMainWindowhWnd_Method3(m_lastSelectedButton.hWnd)
    
    'SetWindowPos m_lastSelectedButton.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    If mainhWnd > 0 Then
        winStyle = GetWindowLong(mainhWnd, GWL_STYLE)

        If (winStyle And WS_MINIMIZE) = WS_MINIMIZE Then
            'ShowWindow mainhWnd, SW_RESTORE
        End If
        
        If Not (winStyle And WS_VISIBLE) = WS_VISIBLE Then
            'ShowWindow mainhWnd, SW_SHOW
        End If
    
        'SetForegroundWindow mainhWnd
    Else
        'mainhWnd = Me.HostForm.hWnd
        'SetForegroundWindow mainhWnd
        'SetActiveWindow mainhWnd
        
        'LogError 0, "MouseDown", "SystemTrayManager", "Couldn't find a master window"
    End If
    
    If Button = vbRightButton Then
        'PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal 0, ByVal WM_NCRBUTTONDOWN
        PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal MAKEWPARAM(cursorPos.X, cursorPos.Y), ByVal WM_RBUTTONDOWN
    ElseIf Button = vbLeftButton Then
        PostMessage ByVal m_lastSelectedButton.hWnd, ByVal m_lastSelectedButton.uCallbackMessage, ByVal MAKEWPARAM(cursorPos.X, cursorPos.Y), ByVal WM_LBUTTONDOWN
        'ActivateAllPopupWindows m_lastSelectedButton.hWnd
    End If
    
    Exit Function

Handler:
    'MsgBox Err.Description
    LogError 0, "MouseDown", "SystemTrayManager", Err.Description
End Function

Private Function GetSelectedItemIndex(ByVal X As Single, ByVal Y As Single) As Long

    Dim yIndex       As Long

    Dim xIndex       As Long

    Dim theItemIndex As Long

    GetSelectedItemIndex = -1

    yIndex = Floor(Y / (MARGIN + ICON_SIZE))
    xIndex = Floor(X / (MARGIN + ICON_SIZE)) + 1

    theItemIndex = xIndex + (yIndex * m_columnLimit)
    
    If theItemIndex > 0 And theItemIndex <= m_lastIconCount Then
        GetSelectedItemIndex = theItemIndex
    End If

End Function

Public Function MouseMove(Position As POINTS)
    
    Dim selectedButton      As TrayButtonInfo

    Dim selectedButtonIndex As Long

    Dim cursorPos           As POINTL

    selectedButtonIndex = GetSelectedItemIndex(Position.X, Position.Y)
    GetCursorPos cursorPos
    
    If selectedButtonIndex <= 0 Or selectedButtonIndex > m_buttons.Count Then

        Exit Function

    End If
    
    Set selectedButton = m_buttons(selectedButtonIndex)

    If selectedButton Is Nothing Then

        Exit Function

    End If

    PostMessage ByVal selectedButton.hWnd, ByVal selectedButton.uCallbackMessage, ByVal MAKEWPARAM(cursorPos.X, cursorPos.Y), ByVal WM_MOUSEMOVE
    
    If m_lastSelectedButton Is selectedButton Then

        Exit Function

    End If
    
    Set m_lastSelectedButton = selectedButton

    If selectedButton Is Nothing Then

        Exit Function

    End If
    
    If selectedButton.Tooltip = vbNullString Then
        
        'm_popup.Hide
    Else
        Debug.Print "selectedButton.Tooltip:: '" & selectedButton.Tooltip & "'"
    
        'm_popup.Top = HostForm.Top + HostForm.Height
        'm_popup.ShowTextPopup selectedButton.Tooltip, Dimensions.Left + (HostForm.Left / Screen.TwipsPerPixelX) + _
         (selectedButtonX) + (SystemTrayHelper.ICON_SIZE / 2)
    End If
                 
End Function

