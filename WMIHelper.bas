Attribute VB_Name = "WMIHelper"
'--------------------------------------------------------------------------------
'    Component  : WMIHelper
'    Project    : ViDock
'
'    Description: All routines that use WMI live here
'                 Windows Management Instrumentation
'
'    Modified   :
'--------------------------------------------------------------------------------

Private m_objWMIService

Private m_initialized As Boolean

Private Function InitializeIfRequired()

    If m_initialized Then Exit Function
    
    Set m_objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    m_initialized = True
End Function

'Very slow
Function GetProcessPath(pId As Long) As String

    On Error GoTo Handler
    
    If pId = 0 Then Exit Function
    
    Dim szExePath   As String

    'Dim szCommandLine As String
    Dim secondQuote As Long

    On Error GoTo Handler

    InitializeIfRequired
    
    Set colItems = m_objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & pId, , 48)
      
    For Each objItem In colItems

        szExePath = objItem.ExecutablePath
    Next
    
    GetProcessPath = szExePath
Handler:
End Function

'Very slow
Public Function GetProcessCommandLineArguments(pId As Long) As String

    On Error GoTo Handler
    
    If pId = 0 Then Exit Function
    
    'Dim szExePath As String
    Dim szCommandLine As String

    Dim secondQuote   As Long

    On Error GoTo Handler

    InitializeIfRequired
    
    Set colItems = m_objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & pId, , 48)
      
    For Each objItem In colItems

        szCommandLine = objItem.CommandLine
    Next
    
    If Len(szCommandLine) < 2 Then
        GetProcessCommandLineArguments = Trim(szCommandLine)

        Exit Function

    End If
    
    secondQuote = InStr(2, szCommandLine, """")

    If secondQuote > 0 Then
        GetProcessCommandLineArguments = TrimNull(Mid$(szCommandLine, secondQuote + 1))

        Exit Function

    End If
        
Handler:
    GetProcessCommandLineArguments = TrimNull(szCommandLine)
End Function
