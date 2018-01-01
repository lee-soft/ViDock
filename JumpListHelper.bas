Attribute VB_Name = "JumpListHelper"
Option Explicit

Private Const EXPLORER_RECENTDOCS         As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"

Private Const EXPLORER_OPENSAVEDOCS_XP    As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSaveMRU"

Private Const EXPLORER_OPENSAVEDOCS_VISTA As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU"

Private EXPLORER_OPENSAVEDOCS             As String

Public Type ShellLink

    szPath As String
    szArguments As String

End Type

Public Function GetMRUListForKey(ByRef srcMRURoot As WinRegistryKey) As String()

    On Error Resume Next

    Dim s_mruList As String

    Dim thisMRU

    Dim thisMRUValue   As String

    Dim thisLnkName    As String

    Dim endFileNamePos As Long

    Dim MRUList()      As String

    Dim lnkFileName    As String

    Dim MRUArrayIndex  As Long

    Debug.Print srcMRURoot.Path

    If srcMRURoot Is Nothing Then Exit Function
    
    s_mruList = srcMRURoot.GetValueAsString("MRUList")

    If srcMRURoot.GetLastError = 0 Then

        While LenB(s_mruList) > 0

            thisMRU = MidB(s_mruList, 1, 2)
            s_mruList = MidB(s_mruList, LenB(thisMRU) + 1)
            
            If LenB(thisMRU) = 2 Then
                thisMRUValue = srcMRURoot.GetValueAsString(CStr(thisMRU))
                lnkFileName = thisMRUValue
                Debug.Print lnkFileName

                If FileExists(lnkFileName) Then
                    ReDim Preserve MRUList(MRUArrayIndex)
                    MRUList(MRUArrayIndex) = lnkFileName
                    
                    MRUArrayIndex = MRUArrayIndex + 1
                End If
            End If

        Wend

    Else
        s_mruList = srcMRURoot.GetValueAsString("MRUListEx")
    
        While LenB(s_mruList) > 0

            thisMRU = MidB(s_mruList, 1, 4)
            
            s_mruList = MidB(s_mruList, LenB(thisMRU) + 1)
            thisMRU = GetDWord(thisMRU)
            
            If thisMRU > -1 Then
                Debug.Print thisMRU
                thisMRUValue = srcMRURoot.GetValueAsString(CStr(thisMRU))
                
                endFileNamePos = 1
                
                'Chr(0) is actually a double byte ZERO ChrB(0) is a single byte
                'Remember strings are double-byte in VB6
                While (Mid(thisMRUValue, endFileNamePos, 1) <> Chr(0)) And endFileNamePos < Len(thisMRUValue)
                    
                    endFileNamePos = endFileNamePos + 1

                Wend
                
                If endFileNamePos > 1 Then
                    thisLnkName = Mid(thisMRUValue, 1, endFileNamePos - 1)

                    If Len(thisLnkName) > 3 Then
                    
                        If Not (Right(thisLnkName, 4) = ".lnk") And InStr(thisLnkName, ".") > 0 Then
                            'lnkFileName = Left(lnkFileName, InStrRev(lnkFileName, ".") - 1) & ".lnk"
                            thisLnkName = thisLnkName & ".lnk"
                        End If
                    End If
                    
                    Debug.Print thisLnkName
                    lnkFileName = GetShortcut(Environ("userprofile") & "\Recent\" & thisLnkName).szPath
                    
                    If FileExists(lnkFileName) Then
                        ReDim Preserve MRUList(MRUArrayIndex)
                        MRUList(MRUArrayIndex) = lnkFileName
                        
                        MRUArrayIndex = MRUArrayIndex + 1
                    End If
                End If
            End If

        Wend

    End If
    
    GetMRUListForKey = MRUList
    
End Function

Function SetOpenSaveDocs()

    Dim thisType As New WinRegistryKey

    thisType.RootKeyType = HKEY_CURRENT_USER
    EXPLORER_OPENSAVEDOCS = EXPLORER_OPENSAVEDOCS_XP
    
    thisType.Path = EXPLORER_OPENSAVEDOCS
    
    If thisType.GetLastError <> 0 Then
        EXPLORER_OPENSAVEDOCS = EXPLORER_OPENSAVEDOCS_VISTA
    End If

End Function

Public Function GetShortcut(lnkSrcFile As String) As ShellLink

    On Error GoTo Handler

    Dim objShell

    Dim objFolder

    Dim sFileName     As String

    Dim sParentFolder As String

    Dim returnData    As ShellLink
    
    sFileName = StrEnd(lnkSrcFile, "\")
    sParentFolder = Left(lnkSrcFile, Len(lnkSrcFile) - Len(sFileName) - 1)

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.NameSpace(sParentFolder & "\")
    
    If (Not objFolder Is Nothing) Then

        Dim objFolderItem
        
        Set objFolderItem = objFolder.ParseName(sFileName)

        If (Not objFolderItem Is Nothing) Then
        
            Dim objShellLink

            Set objShellLink = objFolderItem.GetLink

            If (Not objShellLink Is Nothing) Then
                returnData.szArguments = objShellLink.Arguments
                returnData.szPath = objShellLink.Path

                GetShortcut = returnData
            End If

            Set objShellLink = Nothing
        End If

        Set objFolderItem = Nothing
    End If
    
    Set objFolder = Nothing
    Set objShell = Nothing

    Exit Function

Handler:
End Function

Public Function isLnkTargetValid(sPath As String) As Boolean

    On Error GoTo AssumeYes

    Dim objShell

    Dim objFolder
    
    Dim sFileName     As String

    Dim sParentFolder As String
    
    sFileName = StrEnd(sPath, "\")
    sParentFolder = Left(sPath, Len(sPath) - Len(sFileName) - 1)
    
    If Not (Right(sFileName, 4) = ".lnk") And InStr(sFileName, ".") > 0 Then
        
        sFileName = Left(sFileName, InStrRev(sFileName, ".") - 1) & ".lnk"
    End If
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.NameSpace(sParentFolder & "\")
    
    If (Not objFolder Is Nothing) Then

        Dim objFolderItem
        
        Set objFolderItem = objFolder.ParseName(sFileName)

        If (Not objFolderItem Is Nothing) Then
        
            Dim objShellLink

            Set objShellLink = objFolderItem.GetLink

            If (Not objShellLink Is Nothing) Then
                isLnkTargetValid = FileExists(objShellLink.Path)
            End If

            Set objShellLink = Nothing
        End If

        Set objFolderItem = Nothing
    End If
    
    Set objFolder = Nothing
    Set objShell = Nothing

    Exit Function

AssumeYes:
    isLnkTargetValid = True
    
End Function

Public Function BuildMenuWithoutJumpList() As clsMenu

    Dim srcMenu  As New clsMenu

    srcMenu.AddItem 1, "Minimize"
    srcMenu.AddItem 2, "Restore"
    srcMenu.AddSeperater
    srcMenu.AddItem 5, "Pin"
    srcMenu.AddItem 4, "Close All Windows"

    Set BuildMenuWithoutJumpList = srcMenu
End Function

Public Function BuildMenuWithJumpList(srcCompletePaths)

    Dim thisItem      As String

    Dim srcMenu       As New clsMenu

    Dim jumpListIndex As Long

    Dim jumpListMax   As Long

    If IsArrayInitialized(srcCompletePaths) Then
        jumpListMax = sizeOf(srcCompletePaths)

        If jumpListMax > JUMPLIST_CAP Then
            jumpListMax = JUMPLIST_CAP
        End If
        
        For jumpListIndex = 0 To jumpListMax
            srcMenu.AddItem 7 + jumpListIndex, StrEnd(CStr(srcCompletePaths(jumpListIndex)), "\")
        Next
        
        srcMenu.AddSeperater
    End If

    srcMenu.AddItem 1, "Minimize"
    srcMenu.AddItem 2, "Restore"
    srcMenu.AddSeperater
    srcMenu.AddItem 5, "Pin"
    srcMenu.AddItem 4, "Close All Windows"

    Set BuildMenuWithJumpList = srcMenu

End Function

Public Function GetImageJumpList(ByVal srcImagePath As String) As JumpList

    Dim r_recentDocs   As New WinRegistryKey

    Dim r_openSaveDocs As New WinRegistryKey

    Dim thisType       As WinRegistryKey

    Dim thisImagePath  As String

    Dim setJumpList    As Boolean

    Dim thisJumpList   As New JumpList

    Set GetImageJumpList = thisJumpList
    srcImagePath = UCase(StrEnd(srcImagePath, "\"))
    
    If Len(srcImagePath) = 0 Then Exit Function

    r_openSaveDocs.RootKeyType = HKEY_CURRENT_USER
    r_openSaveDocs.Path = EXPLORER_OPENSAVEDOCS
    
    r_recentDocs.RootKeyType = HKEY_CURRENT_USER
    r_recentDocs.Path = EXPLORER_RECENTDOCS

    thisJumpList.ImageName = srcImagePath
    srcImagePath = UCase(srcImagePath)

    If isset(r_recentDocs.SubKeys) Then

        For Each thisType In r_recentDocs.SubKeys

            thisImagePath = UCase(StrEnd(Trim(GetEXEPathFromQuote(GetAbsolutePath(GetTypeHandlerPath(thisType.Name)))), "\"))
            
            If thisImagePath = srcImagePath Then
                thisJumpList.AddMRURegKey thisType

                setJumpList = True
            End If

        Next

    End If
    
    If isset(r_openSaveDocs.SubKeys) Then

        For Each thisType In r_openSaveDocs.SubKeys

            thisImagePath = UCase(StrEnd(Trim(GetEXEPathFromQuote(GetAbsolutePath(GetTypeHandlerPath("." & thisType.Name)))), "\"))
    
            If thisImagePath = srcImagePath Then
                thisJumpList.AddMRURegKey thisType
                setJumpList = True
            End If

        Next

    End If
    
End Function

Private Function GetTypeHandlerPath(srcType As String)

    Dim thisKey        As New WinRegistryKey

    Dim typeFullName   As String

    Dim primaryCommand As String
    
    thisKey.RootKeyType = HKEY_CLASSES_ROOT
    thisKey.Path = srcType
    
    typeFullName = thisKey.GetValueAsString()
    
    thisKey.Path = typeFullName & "\shell"
    primaryCommand = thisKey.GetValueAsString()

    If primaryCommand = "" Then primaryCommand = "open"
    
    thisKey.Path = typeFullName & "\shell\" & primaryCommand & "\command"
    GetTypeHandlerPath = thisKey.GetValueAsString

    'Debug.Print srcType & "::" & GetTypeHandlerPath

End Function

Public Function GetEXEPathFromQuote(ByVal srcPath As String)

    On Error GoTo Handler
    
    Dim a       As Long

    Dim b       As Long

    a = InStr(srcPath, """") + 1
    b = InStr(a, srcPath, """")
    
    If (a <> 2) Then
        If (a > 1) Then
            'would fetch path in this situation:  C:\blabla\notepad.exe "%1"
            GetEXEPathFromQuote = Trim(Mid(srcPath, 1, a - 2))

            Exit Function

        Else
            'would fetch path in this situation:  C:\blabla\notepad.exe %1
            a = InStr(srcPath, "%") - 1

            If a > 0 Then
                GetEXEPathFromQuote = Left(srcPath, a)
            Else
                GetEXEPathFromQuote = srcPath
            End If
            
            Exit Function

        End If
    End If
    
    If (a > 1 And b > 0 And b > a) Then
        
        GetEXEPathFromQuote = Mid(srcPath, a, (b - a))

        Exit Function

    Else
        a = InStr(srcPath, "%") - 1

        If (a > 0) Then
            GetEXEPathFromQuote = Mid(srcPath, 1, a)

            Exit Function

        End If
    End If
    
    Exit Function

Handler:
    
    GetEXEPathFromQuote = srcPath
End Function

'Replaces all enviromental variables with their absolute equivalents
'It doesn't require that a path be valid either
Public Function GetAbsolutePath(ByVal srcPath As String)
    
    Dim a       As Long

    Dim b       As Long

    Dim varName As String

    Dim spliceA As String

    Dim spliceB As String

    Dim ret     As String

    a = InStr(srcPath, "%") + 1
    b = InStr(a, srcPath, "%")
    
    If (a > 1 And b > 0 And b > a) Then
        
        varName = Mid(srcPath, a, (b - a))
        
        spliceA = Mid(srcPath, 1, a - 2)
        spliceB = Mid(srcPath, b + 1)
        
        ret = spliceA & Environ(varName) & spliceB
    Else
        GetAbsolutePath = srcPath

        Exit Function

    End If
    
    If InStr(ret, "%") > 0 Then
        GetAbsolutePath = GetAbsolutePath(ret)
    Else
        GetAbsolutePath = ret
    End If
    
End Function

