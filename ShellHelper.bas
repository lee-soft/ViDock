Attribute VB_Name = "ShellHelper"
'--------------------------------------------------------------------------------
'    Component  : ShellHelper
'    Project    : ViDock
'
'    Description: Contains functions to help manipulate the OS Shell
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_taskbarhWnd         As Long

Private m_TargetTaskList      As TaskList

Public g_hwndForeGroundWindow As Long

Public Property Get TaskbarHandler()

    Dim temphWnd As Long

    temphWnd = FindWindow("Shell_TrayWnd", "")
    
    If temphWnd <> m_taskbarhWnd Then
        m_taskbarhWnd = temphWnd
    End If

    TaskbarHandler = m_taskbarhWnd
End Property

Public Function EnumWindowsAsTaskList(ByRef srcCollection As TaskList)

    On Error GoTo Handler

    ' Clear list, then fill it with the running
    ' tasks. Return the number of tasks.
    '
    ' The EnumWindows function enumerates all top-level windows
    ' on the screen by passing the handle of each window, in turn,
    ' to an application-defined callback function. EnumWindows
    ' continues until the last top-level window is enumerated or
    ' the callback function returns FALSE.
    '

    If Not srcCollection Is Nothing Then
    
        Set m_TargetTaskList = srcCollection
        Call WinAPIHelper.EnumWindows(AddressOf fEnumWindowsCallBack, ByVal 0)
    End If
    
    Exit Function

Handler:
    LogError Err.Number, "EnumerateWindowsAsTaskObject(); " & Err.Description, "TaskbarHelper"
    
End Function

Public Function fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long

    Dim szWindowClass As String

    szWindowClass = GetWindowClass(hWnd)

    If IsVisibleToTaskBar(hWnd) And Not Settings.ClassBlackList.Exists(szWindowClass) Then
        
        m_TargetTaskList.AddWindowByHwnd hWnd
    End If
    
    fEnumWindowsCallBack = True
End Function

Public Function IsVisibleToTaskBar(hWnd As Long) As Boolean

    Dim lExStyle As Long

    Dim bNoOwner As Boolean

    IsVisibleToTaskBar = False
    
    ' This callback function is called by Windows (from
    ' the EnumWindows API call) for EVERY window that exists.
    ' It populates the listbox with a list of windows that we
    ' are interested in.
    '
    ' Windows to display are those that:
    '   -   are not this app's
    '   -   are visible
    '   -   do not have a parent
    '   -   have no owner and are not Tool windows OR
    '       have an owner and are App windows
    '       can be activated

    If IsWindowVisible(hWnd) Then
        If (getParent(hWnd) = 0) Then
            bNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)

            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
            
                IsVisibleToTaskBar = True
            End If
        End If
    End If

End Function

Public Function IsVisibleOnTaskbar(lExStyle As Long) As Boolean
    IsVisibleOnTaskbar = False

    If (lExStyle And WS_EX_APPWINDOW) Then
        IsVisibleOnTaskbar = True
    End If

End Function

Public Function GetWindowClass(ByVal hWnd As Long) As String

    Dim sClass As String

    sClass = Space$(256)
    GetClassName hWnd, sClass, 255
    GetWindowClass = Left$(sClass, InStr(sClass, vbNullChar) - 1)
End Function
