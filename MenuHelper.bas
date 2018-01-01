Attribute VB_Name = "MenuHelper"
'--------------------------------------------------------------------------------
'    Component  : MenuHelper
'    Project    : ViDock
'
'    Description: [type_description_here]
'
'--------------------------------------------------------------------------------
Option Explicit

Public Function CreateSystemMenu(ByVal hMenu As Long, _
                                 ByVal WindowState As FormWindowStateConstants) As clsMenu

    Dim itemCount    As Long

    Dim objNewMenu   As New clsMenu

    Dim itemIndex    As Long

    Dim itemID       As Long

    Dim bufferString As String

    Dim itemLength   As Long

    Dim itemState    As Long

    Dim menuDefault  As Long

    itemCount = GetMenuItemCount(hMenu)

    For itemIndex = 0 To itemCount - 1
        itemLength = GetMenuString(hMenu, itemIndex, ByVal 0, 0, MF_BYPOSITION) + 1
        bufferString = String(itemLength, 0)
        itemState = GetMenuState(hMenu, itemIndex, MF_BYPOSITION)
        itemID = GetMenuItemID(hMenu, itemIndex)
        GetMenuString hMenu, itemIndex, bufferString, itemLength, MF_BYPOSITION
        
        If WindowState = vbNormal Then
            If itemID = SC_RESTORE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_CLOSE Then
                menuDefault = itemID
            End If
            
        ElseIf WindowState = vbMinimized Then

            If itemID = SC_MOVE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_SIZE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_MINIMIZE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_RESTORE Then
                itemState = MF_STRING
                menuDefault = itemID
            ElseIf itemID = SC_MAXIMIZE Then
                itemState = MF_STRING
            End If
            
        ElseIf WindowState = vbMaximized Then

            If itemID = SC_MAXIMIZE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_MOVE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_SIZE Then
                itemState = MF_GRAYED Or MF_STRING
            End If
        End If
        
        AppendMenu objNewMenu.Handle, itemState, itemID, bufferString
    
        If itemID = menuDefault Then
            SetMenuDefaultItem objNewMenu.Handle, itemIndex, True
        End If

    Next
    
    Set CreateSystemMenu = objNewMenu

End Function

