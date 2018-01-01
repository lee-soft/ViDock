Attribute VB_Name = "MainHelper"
Option Explicit

Public g_DeviceCollection As DeviceCollection

Public g_TaskBar          As TaskBar

Public Settings           As DockSettings

Public SystemMenu         As clsMenu

Public Sub Main()

    GDIPlusCreate (True)
    
    If Not ThemeHelper.Initialize Then
        LogError 0, "Main", "ThemeHelper", "Initialize failed"

        Exit Sub

    End If
    
    'MsgBox "Theme initialized succesfully!", vbInformation

    Set SystemMenu = New clsMenu
    SystemMenu.AddItem 1, "Exit"

    ResizeDesktop

    LogNotice "Starting ViDock rev" & App.Revision
    Set g_DeviceCollection = New DeviceCollection

    Set Settings = New DockSettings
    
    Set g_TaskBar = New TaskBar
    g_TaskBar.Show
    
    'ListMenu.Show
End Sub

Public Function LogNotice(sNotice As String)

    On Error Resume Next
    
    If RunningInVB Then
        Debug.Print Now(), sNotice

        Exit Function

    End If
    
    Dim FileNum As Long

    FileNum = FreeFile
    
    Open App.Path & "\errors.log" For Append As FileNum
    Write #FileNum, Now(), sNotice
    Close FileNum
End Function

Public Function LogError(errNo As Long, _
                         callerFunction As String, _
                         sourceObj As String, _
                         Optional errDescription As String)

    On Error Resume Next
    
    If RunningInVB Then
        Debug.Print Now(), errNo, sourceObj, callerFunction, errDescription
        Exit Function
    End If
    
    Dim FileNum As Long
    FileNum = FreeFile
    
    Open App.Path & "\errors.log" For Append As FileNum
    Write #FileNum, Now(), errNo, sourceObj, callerFunction, errDescription
    Close FileNum
End Function

Public Function SystemMenuHandler(ByVal theResult As Long)
    
    If theResult = 1 Then
        RestoreDesktop

        Unload g_TaskBar
    End If
    
End Function
