Attribute VB_Name = "WinAPIHelper"
'--------------------------------------------------------------------------------
'    Component  : WinAPIHelper
'    Project    : ViDock
'
'    Description: A place to dump Windows API declerations
'
'--------------------------------------------------------------------------------


Option Explicit

Public Const HSHELL_REDRAW As Long = 6

Public Const HSHELL_HIGHBIT = &H8000

Public Const HSHELL_FLASH = 32774

Public Const HSHELL_WINDOWDESTROYED As Long = 2

Public Const HSHELL_WINDOWCREATED   As Long = 1

Public Const HSHELL_WINDOWACTIVATED As Long = 4

Public Const HSHELL_WINDOWREPLACED  As Long = 13

Public Const MSGFLT_ADD = 1

Public Const MSGFLT_REMOVE = 2

Public Const WM_POPUPSYSTEMMENU = &H313

Public Const AW_HOR_POSITIVE = &H1 'Animates the window from left to right. This flag can be used with roll or slide animation.

Public Const AW_HOR_NEGATIVE = &H2 'Animates the window from right to left. This flag can be used with roll or slide animation.

Public Const AW_VER_POSITIVE = &H4 'Animates the window from top to bottom. This flag can be used with roll or slide animation.

Public Const AW_VER_NEGATIVE = &H8 'Animates the window from bottom to top. This flag can be used with roll or slide animation.

Public Const AW_CENTER = &H10 'Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used.

Public Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.

Public Const AW_ACTIVATE = &H20000 'Activates the window.

Public Const AW_SLIDE = &H40000 'Uses slide animation. By default, roll animation is used.

Public Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.

Public Const TME_LEAVE As Long = &H2

Public Const TME_HOVER As Long = &H1

Declare Sub DragAcceptFiles Lib "shell32" (ByVal hWnd As Long, ByVal bool As Long)
Declare Function DragQueryFileW _
        Lib "shell32" (ByVal wParam As Long, _
                       ByVal index As Long, _
                       ByVal lpszFile As Long, _
                       ByVal BufferSize As Long) As Long
Declare Sub DragFinish Lib "shell32" (ByVal hDrop As Integer)

'Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TrackMouseEvent _
               Lib "user32.dll" (ByRef lpEventTrack As TrackMouseEvent) As Long

Public Declare Function SetMenuDefaultItem _
               Lib "user32.dll" (ByVal hMenu As Long, _
                                 ByVal uItem As Long, _
                                 ByVal fByPos As Long) As Long
Declare Function ChangeWindowMessageFilter _
        Lib "user32.dll" (ByVal Message As Long, _
                          ByVal dwFlag As Integer) As Boolean

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpszOp As String, _
                                      ByVal lpszFile As String, _
                                      ByVal lpszParams As String, _
                                      ByVal LpszDir As String, _
                                      ByVal FsShowCmd As Long) As Long
                    
Public Declare Function AnimateWindow _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal dwTime As Long, _
                             ByVal dwFlags As Long) As Boolean

Public Declare Function IIDFromString _
               Lib "ole32.dll" (ByVal lpsz As Long, _
                                ByRef lpiid As Any) As Long

Public Declare Function SHGetImageListXP _
               Lib "shell32.dll" _
               Alias "#727" (ByVal iImageList As Long, _
                             ByRef riid As Long, _
                             ByRef ppv As Any) As Long

Public Declare Function SHGetImageList _
               Lib "shell32.dll" (ByVal iImageList As Long, _
                                  ByRef riid As Long, _
                                  ByRef ppv As Any) As Long

Public Declare Function EnumWindows _
               Lib "user32.dll" (ByVal lpEnumFunc As Long, _
                                 ByVal lParam As Long) As Long

Public Declare Function CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (ByRef pDest As Any, _
                                      ByRef pSource As Any, _
                                      ByVal dwLength As Long) As Long

Public Declare Function SendMessageTimeout _
               Lib "user32" _
               Alias "SendMessageTimeoutA" (ByVal hWnd As Long, _
                                            ByVal msg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As Long, _
                                            ByVal fuFlags As Long, _
                                            ByVal uTimeout As Long, _
                                            lpdwResult As Long) As Long

Public Declare Function PrintWindow _
               Lib "user32.dll" (ByVal hWnd As Long, _
                                 ByVal hdcBlt As Long, _
                                 ByVal nFlags As Long) As Long

Public Declare Function RedrawWindow _
               Lib "user32" (ByVal hWnd As Long, _
                             lprcUpdate As Any, _
                             ByVal hrgnUpdate As Long, _
                             ByVal fuRedraw As Long) As Long

Public Declare Function Wow64DisableWow64FsRedirection _
               Lib "kernel32.dll" (ByRef oldValue As Long) As Long

Public Declare Function Wow64RevertWow64FsRedirection _
               Lib "kernel32.dll" (ByRef oldValue As Long) As Long

Public Declare Function GetSystemWow64Directory _
               Lib "kernel32.dll" _
               Alias "GetSystemWow64DirectoryA" (ByVal lpBuffer As String, _
                                                 ByVal uSize As Long) As Integer

Public Declare Function SHGetFileInfo _
               Lib "shell32.dll" _
               Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                       ByVal dwFileAttributes As Long, _
                                       psfi As SHFILEINFO, _
                                       ByVal cbFileInfo As Long, _
                                       ByVal uFlags As Long) As Long

Public Declare Function GetModuleFileNameExW _
               Lib "psapi.dll" (ByVal hProcess As Long, _
                                ByVal hModule As Long, _
                                ByVal lpFileName As Long, _
                                ByVal nSize As Long) As Long

Public Declare Function GetProcessImageFileName _
               Lib "psapi" _
               Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, _
                                                 ByVal lptrImageFileName As Long, _
                                                 ByVal nSize As Long) As Long

Public Declare Function GetProcAddress _
               Lib "kernel32" (ByVal hModule As Long, _
                               ByVal lpProcName As String) As Long

Public Declare Function IsWow64Process _
               Lib "kernel32" (ByVal hProc As Long, _
                               bWow64Process As Boolean) As Long
    
Public Declare Function RegisterShellHookWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function RegisterWindowMessage _
               Lib "user32" _
               Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Public Const SMTO_BLOCK = &H1

Public Const SMTO_ABORTIFHUNG = &H2

Public Const RDW_ALLCHILDREN = &H80

Public Const RDW_ERASE = &H4

Public Const RDW_INVALIDATE = &H1

Public Const RDW_UPDATENOW = &H100

Public Const RDW_INTERNALPAINT = &H2

Public Const RDW_VALIDATE = &H8
