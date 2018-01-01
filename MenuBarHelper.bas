Attribute VB_Name = "MenuBarHelper"
'--------------------------------------------------------------------------------
'    Component  : MenuBarHelper
'    Project    : ViDock
'
'    Description: Containts MenuBar helper functions
'
'--------------------------------------------------------------------------------
Option Explicit

Public Const TEXTMODE_ITEM_Y_GAP  As Long = 21

Public Const ITEM_MARGIN_X        As Long = 20

Public Const MENU_MARGIN_X        As Long = 5

Public Const MENU_MARGIN_Y        As Long = 2

Public Const MAX_TITLE_CHARACTERS As Long = 30

Public Function HandleProcessMenuResult(theMenuResult As Long, lastProcess As Process)

    Dim szNewShortcutPath     As String

    Dim szOpenedFileInProcess As String

    Select Case theMenuResult
    
        Case 1
            lastProcess.MinimizeAllWindows
    
        Case 2
            lastProcess.RestoreAllWindows
        
        Case 4
            lastProcess.RequestCloseAllWindows
        
        Case 5

            If lastProcess.Pinned = False Then
                If lastProcess.Path <> vbNullString Then
                
                    If Len(lastProcess.Arguments) > 0 Then
                        szOpenedFileInProcess = StrEnd(lastProcess.Arguments, "\")
                        szNewShortcutPath = Environ("appdata") & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\" & lastProcess.Caption & " - " & szOpenedFileInProcess & ".lnk"
                    Else
                        szNewShortcutPath = Environ("appdata") & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\" & lastProcess.Caption & ".lnk"
                
                    End If
                
                    If CreateShortcut(szNewShortcutPath, lastProcess.Path, lastProcess.Arguments) Then
                    
                        lastProcess.PhysicalLinkFile = szNewShortcutPath
                        lastProcess.Pinned = True
                    End If
                
                Else
                    LogError -2, "Invalid path; Null", "TaskBar"
                End If

            Else
                lastProcess.Pinned = False
            
                If FileExists(lastProcess.PhysicalLinkFile) Then

                    On Error GoTo Handler

                    Kill lastProcess.PhysicalLinkFile
Handler:
                End If
            
                'ViGlance
                'SortPinnedList
                'DumpOptions
            
                'RemoveFromPinnedList m_lastHoveredGroup.Path
            End If

    End Select

End Function

Public Function GetWindowTextByhWnd(ByVal hWnd As Long) As String

    Dim lReturn     As Long

    Dim windowTitle As String

    windowTitle = Space$(256)
    
    lReturn = GetWindowText(hWnd, windowTitle, Len(windowTitle))

    If lReturn Then
        windowTitle = Left$(windowTitle, lReturn)
        
        If Len(windowTitle) > MAX_TITLE_CHARACTERS Then
            windowTitle = Left(windowTitle, MAX_TITLE_CHARACTERS) & "... "
        End If
    End If

    GetWindowTextByhWnd = windowTitle
End Function

Private Function HandleSubMenu(ByRef theMenuItem As MenuItem, _
                               ByRef sourceMenu As ListMenu, _
                               ByVal theItemPosition As Long)

    If Not sourceMenu.ChildMenu Is Nothing Then
        Unload sourceMenu.ChildMenu
    End If
    
    Set sourceMenu.ChildMenu = New ListMenu
    Set sourceMenu.ChildMenu.ParentMenu = sourceMenu
    
    sourceMenu.ChildMenu.ShowList theMenuItem.Children, (sourceMenu.Top + ((theItemPosition * TEXTMODE_ITEM_Y_GAP) * Screen.TwipsPerPixelY)) - 45 * Screen.TwipsPerPixelY, (sourceMenu.Left / Screen.TwipsPerPixelX) + sourceMenu.ScaleWidth, vbPopupMenuLeftAlign, False, True
End Function

Public Function HandleMenuItemHovered(ByRef theMenuItem As MenuItem, _
                                      ByRef sourceMenu As ListMenu, _
                                      ByVal theItemPosition As Long, _
                                      ByRef closeCallerMenu As Boolean)

    If theMenuItem.Children.Count > 0 Then
        HandleSubMenu theMenuItem, sourceMenu, theItemPosition
        closeCallerMenu = False

        Exit Function

    Else

        If Not sourceMenu.ChildMenu Is Nothing Then
            sourceMenu.ChildMenu.closeMe
            Set sourceMenu.ChildMenu = Nothing
        End If
    End If

End Function

Public Function HandleMenuItem(ByRef theMenuItem As MenuItem, _
                               ByRef sourceMenu As ListMenu, _
                               ByVal theItemPosition As Long, _
                               ByRef closeCallerMenu As Boolean)

    If theMenuItem.Children.Count > 0 Then
        HandleSubMenu theMenuItem, sourceMenu, theItemPosition
        closeCallerMenu = False

        Exit Function

    End If
    
    PostMessage theMenuItem.hWnd, ByVal WM_COMMAND, ByVal theMenuItem.itemID, ByVal 0
    closeCallerMenu = True
End Function

Public Function PopulateMenuFromHandle(ByRef theMenuRoot As Collection, _
                                       ByVal hMenu As Long, _
                                       ByVal hWnd As Long)

    Dim num           As Long

    Dim itemIndex     As Long

    Dim thisMenuItem  As MenuItem

    Dim hSubMenu      As Long

    Dim szCaption     As String

    Dim captionLength As Long

    Dim tabPosition   As Long

    num = GetMenuItemCount(hMenu)

    If num > -1 Then Debug.Print "MenuCount:: " & num
    
    For itemIndex = 0 To num - 1
        Set thisMenuItem = New MenuItem
        hSubMenu = GetSubMenu(hMenu, itemIndex)
        
        szCaption = Space(256)
        captionLength = GetMenuString(hMenu, itemIndex, szCaption, Len(szCaption), MF_BYPOSITION)
        szCaption = Replace(Left$(szCaption, captionLength), "&", "")
        
        tabPosition = InStrRev(szCaption, vbTab)

        If tabPosition > 0 Then
            szCaption = Mid(szCaption, 1, tabPosition)
        End If
        
        thisMenuItem.Caption = szCaption
        thisMenuItem.itemID = GetMenuItemID(hMenu, itemIndex)
        thisMenuItem.hWnd = hWnd

        theMenuRoot.Add thisMenuItem
        PopulateMenuFromHandle thisMenuItem.Children, hSubMenu, hWnd
    Next

End Function

Private Sub GetMenuInfo(hMenu As Long, spaces As Integer, txt As String)

    Dim num       As Integer

    Dim i         As Integer

    Dim length    As Long

    Dim sub_hmenu As Long

    Dim sub_name  As String

    num = GetMenuItemCount(hMenu)

    For i = 0 To num - 1
        ' Save this menu's info.
        sub_hmenu = GetSubMenu(hMenu, i)
        Debug.Print GetMenuItemID(hMenu, i)

        sub_name = Space$(256)
        length = GetMenuString(hMenu, i, sub_name, Len(sub_name), MF_BYPOSITION)

        sub_name = Left$(sub_name, length)

        txt = txt & Space$(spaces) & sub_name & vbCrLf

        ' Get its child menu's names.
        GetMenuInfo sub_hmenu, spaces + 4, txt
    Next i

End Sub
