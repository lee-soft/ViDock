Attribute VB_Name = "DesktopHelper"
'--------------------------------------------------------------------------------
'    Component  : DesktopHelper
'    Project    : ViDock
'
'    Description: Changes the workspace area on the desktop
'
'--------------------------------------------------------------------------------
Option Explicit

Public Function RestoreDesktop()

    Dim originalDesktop As Win.RECT

    originalDesktop = GetWorkspace
    
    originalDesktop.Top = originalDesktop.Top - 24

    Call SetWorkspace(originalDesktop)
End Function

Public Function ResizeDesktop()
    
    Dim originalDesktop As Win.RECT

    originalDesktop = GetWorkspace
    
    originalDesktop.Top = originalDesktop.Top + 24

    Call SetWorkspace(originalDesktop)
    
End Function

Public Function GetWorkspace() As Win.RECT

    Dim result    As Long

    Dim workspace As Win.RECT

    result = SystemParametersInfo(SPI_GETWORKAREA, vbNull, workspace, 0)

    If result = 0 Then
        LogError GetLastError(), "GetWorkspace", "DesktopHelper", "failed to get the user workspace"

        Exit Function

    End If
    
    GetWorkspace = workspace
End Function

Public Function SetWorkspace(ByRef newSpace As Win.RECT) As Boolean

    Dim result As Long

    result = SystemParametersInfo(SPI_SETWORKAREA, vbNull, newSpace, SPIF_UPDATEINIFILE)
    
    If result = 0 Then
        LogError GetLastError(), "SetWorkspace", "DesktopHelper", "failed to set the user workspace"

        Exit Function

    End If
    
    SetWorkspace = True
End Function
