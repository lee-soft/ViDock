Attribute VB_Name = "SystemTrayHelper"
'--------------------------------------------------------------------------------
'    Component  : SystemTrayHelper
'    Project    : ViDock
'
'    Description: Contains procedures to assist with Window's system tray
'                 enumeration
'
'--------------------------------------------------------------------------------
Option Explicit

Public Const ICON_SIZE As Long = 16

Public Const MARGIN    As Long = 2

Private Declare Function GetProcAddress _
                Lib "kernel32" (ByVal hModule As Long, _
                                ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle _
                Lib "kernel32" _
                Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function IsWow64Process _
                Lib "kernel32" (ByVal hProc As Long, _
                                bWow64Process As Boolean) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function FindWindowEx _
                Lib "user32" _
                Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                       ByVal hWnd2 As Long, _
                                       ByVal lpsz1 As String, _
                                       ByVal lpsz2 As String) As Long

Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hWnd As Long, _
                              lpdwProcessId As Long) As Long

Private Declare Function OpenProcess _
                Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                ByVal bInheritHandle As Long, _
                                ByVal dwProcessId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function VirtualFreeEx _
                Lib "kernel32" (ByVal hProcess As Long, _
                                lpAddress As Any, _
                                ByVal dwSize As Long, _
                                ByVal dwFreeType As Long) As Long

Private Declare Function VirtualAllocEx _
                Lib "kernel32" (ByVal hProcess As Long, _
                                lpAddress As Any, _
                                ByVal dwSize As Long, _
                                ByVal flAllocationType As Long, _
                                ByVal flProtect As Long) As Long

Private Declare Function ReadProcessMemory _
                Lib "kernel32" (ByVal hProcess As Long, _
                                lpBaseAddress As Any, _
                                lpBuffer As Any, _
                                ByVal nSize As Long, _
                                Optional lpNumberOfBytesWritten As Long) As Long
                                

Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long

Private Declare Function EnumThreadWindows _
                Lib "user32.dll" (ByVal dwThreadId As Long, _
                                  ByVal lpfn As Long, _
                                  ByVal lParam As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Const TB_GETITEMRECT = 1053

Const MEM_COMMIT = &H1000

Const MEM_RESERVE = &H2000

Const MEM_RELEASE = &H8000&

Const PAGE_READWRITE = &H4

Const MAX_PATH = 260

Const WM_USER = &H400

Const TB_BUTTONCOUNT = (WM_USER + 24)

Const TB_GETBUTTON = (WM_USER + 23)

Const TB_GETBUTTONTEXTA = (WM_USER + 45)

Const TBSTATE_HIDDEN = 8

Public Const PROCESS_QUERY_INFORMATION = (&H400)

Const PROCESS_VM_READ = (&H10)

Const PROCESS_VM_WRITE = (&H20)

Const PROCESS_VM_OPERATION = (&H8)

Const ACCESS = PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION

Const SIZEOF_INT64 As Long = 8

Public Type PROCESSENTRY32

    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH

End Type

Public Type TBBUTTON

    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    Reserved(1) As Byte
    dwData As Long
    iString As Long

End Type

Public Type TBBUTTON_64

    iBitmap      As Long
    idCommand    As Long
    fsState      As Byte
    fsStyle      As Byte
    Reserved(5) As Byte
  
    dwData       As Long
    dwDataPadding As Long 'uint64 - the extra 4 bytes that are probably never going to be used
  
    iString      As Long
    iStringPadding As Long 'uint64 - the extra 4 bytes that are probably never going to be used

End Type

Public Type TRAYDATA

    hWnd  As Long
    uID As Long
    uCallbackMessage As Long
    bReserved(2) As Byte
    hIcon As Long

End Type

Public Type TRAYDATA_64

    hWnd As Long
    hWndPadding As Long 'hWnd uint64 - the extra 4 bytes are NEVER used
    
    uID As Long
    uCallbackMessage As Long
    bReserved(4) As Byte
    
    hIcon As Long
    hIconPadding As Long 'hIcon uint64 - the extra 4 bytes are NEVER used

End Type

Public Type Int64 'LowPart must be the first one in LittleEndian systems

    'required part
    LowPart  As Long
    HighPart As Long

    'optional part
    SignBit As Byte 'define this as long if you want to get minimum CPU access time.

End Type

Public Type HANDLE_DATA

    ProcessID As Long
    BestHandle As Long

End Type

Private m_returnHandleData   As HANDLE_DATA

Private m_dirtySearch        As Boolean

Public Function ActivateAllPopupWindows(ByVal hWnd As Long)

    'Dim pId As Long
    'Dim tId As Long

    '    tId = GetWindowThreadProcessId(hWnd, pId)
    '    Call EnumThreadWindows(tId, AddressOf fEnumActivatePopupWindowThreadWindowsCallback, 0)

    Dim mainhWnd As Long

    mainhWnd = GetMainWindowhWnd_Method3(hWnd)
    
    ShowWindow mainhWnd, SW_SHOW
    SetForegroundWindow mainhWnd
    SetActiveWindow mainhWnd
End Function

Public Function GetMainWindowFromhWnd(ByVal hWnd As Long) As Long

    Dim nRet      As Long

    Dim nMainhWnd As Long

    If HasNoOwner(hWnd) Then
        GetMainWindowFromhWnd = hWnd

        Exit Function

    End If

    nRet = GetWindowLong(hWnd, GWL_HWNDPARENT)

    Do While nRet
        nMainhWnd = nRet
        nRet = GetWindowLong(nMainhWnd, GWL_HWNDPARENT)
    Loop

    GetMainWindowFromhWnd = nMainhWnd
End Function

Public Function GetMainWindowhWnd_Method3(ByVal hWnd As Long) As Long
    
    Dim pId As Long

    Dim tId As Long

    tId = GetWindowThreadProcessId(hWnd, pId)
    m_returnHandleData.BestHandle = 0
    m_dirtySearch = False
    
    Call EnumThreadWindows(tId, AddressOf fEnumThreadWindowsCallback, 0)
    GetMainWindowhWnd_Method3 = m_returnHandleData.BestHandle
    
    If GetMainWindowhWnd_Method3 = 0 Then
        m_dirtySearch = True
        
        Call EnumThreadWindows(tId, AddressOf fEnumThreadWindowsCallback, 0)
        GetMainWindowhWnd_Method3 = m_returnHandleData.BestHandle
    End If

End Function

Public Function fEnumThreadWindowsCallback(ByVal hWnd As Long, _
                                           ByVal lParam As Long) As Long

    If m_dirtySearch Then
        If HasNoOwner(hWnd) Then
    
            m_returnHandleData.BestHandle = hWnd
            fEnumThreadWindowsCallback = False

            Exit Function

        End If

    Else

        If IsMainWindow(hWnd) Then
    
            m_returnHandleData.BestHandle = hWnd
            fEnumThreadWindowsCallback = False

            Exit Function

        End If
    End If

    fEnumThreadWindowsCallback = True
End Function

Public Function fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long

    Dim pId  As Long

    Dim data As HANDLE_DATA
    
    'convert pointer back to data
    WinAPIHelper.CopyMemory data, lParam, LenB(data)
    GetWindowThreadProcessId hWnd, pId

    If data.ProcessID = pId Then
        If HasNoOwner(hWnd) Then

            m_returnHandleData.BestHandle = hWnd
            fEnumWindowsCallBack = False

            Exit Function

        End If
    End If

    fEnumWindowsCallBack = True
End Function

Public Function fEnumActivatePopupWindowThreadWindowsCallback(ByVal hWnd As Long, _
                                                              ByVal lParam As Long) As Long
    Debug.Print "fEnumActivatePopupWindowThreadWindowsCallback: hWnd"

    Dim winStyle As Long
    
    winStyle = GetWindowLong(hWnd, GWL_STYLE)

    If (winStyle And WS_POPUP) = WS_POPUP Then
        Debug.Print "Activating: " & hWnd
        ShowWindow hWnd, SW_RESTORE
        ShowWindow hWnd, SW_SHOW
        SetActiveWindow hWnd
        SetForegroundWindow hWnd
        
        'fEnumActivatePopupWindowThreadWindowsCallback = False
        'Exit Function
    End If

    fEnumActivatePopupWindowThreadWindowsCallback = True
End Function

Public Function IsMainWindow(hWnd As Long) As Boolean
    IsMainWindow = GetWindow(hWnd, GW_OWNER) = 0 And IsWindowVisible(hWnd)
End Function

Public Function HasNoOwner(hWnd As Long) As Boolean
    HasNoOwner = GetWindow(hWnd, GW_OWNER) = 0
End Function

Public Function FindTrayToolbarWindowPromotedIcons() As Long
    '; User Promoted Notification Area
    
    Dim trayHandle As Long
    
    trayHandle = FindWindow("Shell_TrayWnd", vbNullString)

    If trayHandle = 0 Then
        LogError 0, "FindTrayToolbarWindowPromotedIcons", "SystemTrayHelper", "Invalid TrayhWnd"

        Exit Function

    End If
    
    trayHandle = FindWindowEx(trayHandle, 0, "TrayNotifyWnd", vbNullString)

    If trayHandle = 0 Then
        LogError 0, "FindTrayToolbarWindowPromotedIcons", "SystemTrayHelper", "Invalid TrayNotityhWnd"

        Exit Function

    End If
    
    trayHandle = FindWindowEx(trayHandle, 0, "SysPager", vbNullString)

    If trayHandle = 0 Then
        LogError 0, "FindTrayToolbarWindowPromotedIcons", "SystemTrayHelper", "Invalid SysPagerhWnd"

        Exit Function

    End If

    trayHandle = FindWindowEx(trayHandle, 0, "ToolbarWindow32", vbNullString)

    If trayHandle = 0 Then
        LogError 0, "FindTrayToolbarWindowPromotedIcons", "SystemTrayHelper", "Invalid ToolbarWindow32hWnd"

        Exit Function

    End If

    MsgBox trayHandle
    FindTrayToolbarWindowPromotedIcons = trayHandle
End Function

Public Function FindTrayToolbarWindowHiddenIcons() As Long
    '; NotifyIconOverflowWindow for Windows 7
    
    Dim iconOverflowHandle As Long

    Dim trayToolbarHandle  As Long

    iconOverflowHandle = FindWindow("NotifyIconOverflowWindow", vbNullString)
    
    '; Overflow Notification Area
    trayToolbarHandle = FindWindowEx(iconOverflowHandle, 0, "ToolbarWindow32", vbNullString)
    
    FindTrayToolbarWindowHiddenIcons = trayToolbarHandle
End Function

Public Function SysTrayIconCount(trayhWnd As Long) As Long

    Dim iconCount As Long

    If trayhWnd = -1 Then Exit Function
    iconCount = SendMessage(trayhWnd, TB_BUTTONCOUNT, ByVal 0, ByVal 0)
    
    SysTrayIconCount = iconCount
End Function

Public Function SysTrayGetButtonInfo_86(trayhWnd As Long, _
                                        pId As Long, _
                                        iIndex As Long) As TrayButtonInfo

    Dim ret             As Long

    Dim lpData          As Long

    Dim td              As TRAYDATA

    Dim tbb             As TBBUTTON

    Dim hProcess        As Long

    Dim baToolTip(1024) As Byte

    Dim szToolTip       As String

    Dim returnInfo      As TrayButtonInfo

    Set returnInfo = New TrayButtonInfo
    Set SysTrayGetButtonInfo_86 = returnInfo

    If trayhWnd = -1 Then Exit Function
    hProcess = OpenProcess(ACCESS, 0, pId)

    If hProcess = 0 Then Exit Function
    
    'lpData = VirtualAllocEx(hProcess, 0&, LenB(tbb), MEM_COMMIT, PAGE_READWRITE)
    'If lpData = 0 Then
    lpData = VirtualAllocEx(hProcess, ByVal 0&, LenB(tbb), MEM_COMMIT, PAGE_READWRITE)
    'If lpData = 0 Then
    'lpData = VirtualAllocEx2(hProcess, ByVal 0&, LenB(tbb), MEM_COMMIT, PAGE_READWRITE)
            
    'If Not m_nullLpDataNotified Then
    'LogError GetLastError(), "SysTrayGetButtonInfo_X86", "VirtualAllocEx Failed to allocate memory in the normal way (mem error?) attempting alternate VirtualAllocEx decleration"
    'm_nullLpDataNotified = True
    'End If
    
    'End If
    'End If

    If lpData = 0 Then
        LogError GetLastError(), "SysTrayGetButtonInfo_X86", "VirtualAllocEx Failed to allocate memory"

        CloseHandle hProcess

        Exit Function

    End If
    
    ret = SendMessage(trayhWnd, TB_GETBUTTON, ByVal iIndex, ByVal lpData)

    If ret <> 0 Then
        ReadProcessMemory hProcess, ByVal lpData, tbb, LenB(tbb)
        ReadProcessMemory hProcess, ByVal tbb.dwData, td, Len(td)
        
        ReadProcessMemory hProcess, ByVal tbb.iString, baToolTip(0), UBound(baToolTip), 0
        szToolTip = CStr(baToolTip)
        szToolTip = Left(szToolTip, InStr(szToolTip, vbNullChar) - 1)
        
        If (tbb.fsState Xor TBSTATE_HIDDEN) Then
            returnInfo.Visible = True
            
            'ret = SendMessage(trayHwnd, TB_GETITEMRECT, ByVal iIndex, ByVal lpData)
            'ReadProcessMemory hProcess, ByVal lpData, ByVal tbItemRect, Len(tbItemRect), 0
        End If
        
        returnInfo.Tooltip = szToolTip
        returnInfo.hIcon = td.hIcon
        returnInfo.hWnd = td.hWnd
        returnInfo.uCallbackMessage = td.uCallbackMessage
        returnInfo.uID = td.uID
    End If

    If VirtualFreeEx(hProcess, ByVal lpData, 0, MEM_RELEASE) = 0 Then
        LogError GetLastError(), "SysTrayGetButtonInfo_X86", "SysTrayHelper", "Failed to release memory"
    End If
    
    CloseHandle hProcess
    
End Function

Public Function SysTrayGetButtonInfo_64(trayhWnd As Long, _
                                        pId As Long, _
                                        iIndex As Long) As TrayButtonInfo

    Dim ret             As Long

    Dim lpData          As Long

    Dim td              As TRAYDATA_64

    Dim tbb             As TBBUTTON_64

    Dim hProcess        As Long

    Dim baToolTip(1024) As Byte

    Dim szToolTip       As String

    Dim returnInfo      As TrayButtonInfo

    Set returnInfo = New TrayButtonInfo
    Set SysTrayGetButtonInfo_64 = returnInfo

    If trayhWnd = -1 Then Exit Function
    hProcess = OpenProcess(ACCESS, 0, pId)

    If hProcess = 0 Then Exit Function

    'lpData = VirtualAllocEx(hProcess, 0&, LenB(tbb), MEM_COMMIT, PAGE_READWRITE)
    'If lpData = 0 Then
    lpData = VirtualAllocEx(hProcess, ByVal 0&, LenB(tbb), MEM_COMMIT, PAGE_READWRITE)
    '    If lpData = 0 Then
    '        lpData = VirtualAllocEx2(hProcess, ByVal 0&, LenB(tbb), MEM_COMMIT, PAGE_READWRITE)
    '
    '        If Not m_nullLpDataNotified Then
    '            LogError GetLastError(), "SysTrayGetButtonInfo_X64", "VirtualAllocEx Failed to allocate memory in the normal way (mem error?) attempting alternate VirtualAllocEx decleration"
    '            m_nullLpDataNotified = True
    '        End If
    '    End If
    'End If

    If lpData = 0 Then
        LogError GetLastError(), "SysTrayGetButtonInfo_X64", "VirtualAllocEx Failed to allocate memory"

        CloseHandle hProcess

        Exit Function

    End If
    
    If lpData <> 0 Then
        ret = SendMessage(trayhWnd, TB_GETBUTTON, ByVal iIndex, ByVal lpData)

        If ret <> 0 Then
            ReadProcessMemory hProcess, ByVal lpData, tbb, LenB(tbb)
            ReadProcessMemory hProcess, ByVal tbb.dwData, td, Len(td)
            
            ReadProcessMemory hProcess, ByVal tbb.iString, baToolTip(0), UBound(baToolTip), 0
            szToolTip = CStr(baToolTip)
            szToolTip = Left(szToolTip, InStr(szToolTip, vbNullChar) - 1)
            
            'Debug.Print "szToolTip: " & szToolTip
            
            If (tbb.fsState Xor TBSTATE_HIDDEN) Then
                returnInfo.Visible = True
                
                'ret = SendMessage(trayHwnd, TB_GETITEMRECT, ByVal iIndex, ByVal lpData)
                'ReadProcessMemory hProcess, ByVal lpData, ByVal tbItemRect, Len(tbItemRect), 0
            End If
            
            returnInfo.Tooltip = szToolTip
            returnInfo.hIcon = td.hIcon
            returnInfo.hWnd = td.hWnd
            returnInfo.uCallbackMessage = td.uCallbackMessage
            returnInfo.uID = td.uID
        End If
    End If
    
    If VirtualFreeEx(hProcess, ByVal lpData, 0, MEM_RELEASE) = 0 Then
        LogError GetLastError(), "SysTrayGetButtonInfo_X64", "SysTrayHelper", "Failed to release memory"
    End If

    'If VirtualFreeEx(hProcess, ByVal lpData, ByVal LenB(tbb), ByVal MEM_DECOMMIT) = 0 Then
    'LogError 0, "SysTrayGetButtonInfo_X64", "SysTrayHelper", "Failed to decommit memory"
    'End If
    
    CloseHandle hProcess
    
End Function

Public Function IsProcess64bit(hProcess As Long) As Boolean

    Dim Handle As Long, runningAsWow64 As Boolean

    ' Assume initially that this is a Wow64 process
    runningAsWow64 = True

    ' Now check to see if IsWow64Process function exists
    Handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    If Handle > 0 Then ' IsWow64Process function exists
    
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process hProcess, runningAsWow64
    End If
    
    If runningAsWow64 Then
        IsProcess64bit = False
    Else
        IsProcess64bit = True
    End If

End Function

Public Function GetSelectedTrayButton(ByRef trayButtons As Collection, _
                                      ByRef buttonX As Long, _
                                      ByVal X As Single) As TrayButtonInfo

    Dim thisItem As TrayButtonInfo

    Dim startX   As Long

    Dim endX     As Long

    startX = 0

    For Each thisItem In trayButtons

        'If thisItem.Visible Then
        endX = startX + (SystemTrayHelper.ICON_SIZE + SystemTrayHelper.MARGIN)
            
        If X > startX And X < endX Then
            buttonX = startX
            Set GetSelectedTrayButton = thisItem

            Exit For

        End If
            
        startX = endX
        'End If
    Next

End Function

