Attribute VB_Name = "ProcessHelper"
'--------------------------------------------------------------------------------
'    Component  : ProcessHelper
'    Project    : ViDock
'
'    Description: Contains functions to help manipulate processes in the OS
'
'--------------------------------------------------------------------------------
Option Explicit

Public Function GetProcessID(hWnd As Long) As Long

    Dim Id As Long

    GetWindowThreadProcessId hWnd, Id
    GetProcessID = Id
End Function

Function GetProcessPath(pId As Long)

    Dim lReturn         As Long

    Dim szFileName      As String

    Dim Buffer(2048)    As Byte

    Dim dwLength        As Long

    lReturn = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pId)

    If lReturn = 0 Then
        Debug.Print "Failed to open process: " & pId

        Exit Function

    End If
        
    dwLength = GetModuleFileNameExW(lReturn, 0, VarPtr(Buffer(0)), UBound(Buffer))

    If dwLength = 0 Then
        dwLength = GetProcessImageFileName(lReturn, VarPtr(Buffer(0)), UBound(Buffer))
        
        szFileName = Left$(Buffer, dwLength)
        szFileName = g_DeviceCollection.ConvertToLetterPath(szFileName)

        If dwLength = 0 Then
            CloseHandle lReturn

            Exit Function

        End If

    Else
        szFileName = Left$(Buffer, dwLength)
    End If
    
    If Is64bit Then
        If InStr(szFileName, "SysWOW64") > 0 Then
            szFileName = Replace(szFileName, "SysWOW64", "System32")
        End If
    End If
    
    GetProcessPath = szFileName
    
    CloseHandle lReturn
End Function

Public Function Is64bit() As Boolean

    Dim DirPath As String * 255
    
    Dim result  As Long

    result = GetSystemWow64Directory(DirPath, 255)
    
    If result > 0 Then
        Is64bit = True
    End If

End Function

Public Function IsPIDValid(ByVal pId As Long) As Long

    Debug.Print "IsPIDValid():PID:: " & IsPIDValid

    Dim hProcess As Long

    Dim dwRetval As Long
    
    If (pId = 0) Then
        IsPIDValid = 1: Exit Function
    End If
    
    If (pId < 0) Then
        IsPIDValid = 0: Exit Function
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pId)

    If (hProcess = pNull) Then

        'invalid parameter means PID isn't in the system
        If (WinBase.GetLastError() = ERROR_INVALID_PARAMETER) Then
            IsPIDValid = 0: Exit Function
        End If

        'some other error
        IsPIDValid = -1
    End If

    dwRetval = WaitForSingleObject(hProcess, 0)
    Call CloseHandle(hProcess)  'otherwise you'll be losing handles
    
    Select Case dwRetval
    
        Case WAIT_OBJECT_0:
            IsPIDValid = 0: Exit Function

        Case WAIT_TIMEOUT:
            IsPIDValid = 1: Exit Function
        
        Case Else
            IsPIDValid = -1: Exit Function

    End Select

End Function

