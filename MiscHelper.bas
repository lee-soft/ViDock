Attribute VB_Name = "MiscHelper"
'--------------------------------------------------------------------------------
'    Component  : MiscHelper
'    Project    : ViDock
'
'    Description: Functions that had nowhere else to live :(
'                 TODO: Create dedicated modules for each function no matter
'                 how slim they might be
'
'--------------------------------------------------------------------------------
Option Explicit

Public pngIndex As Long

Public Function CreateShortcut(szShortcutPath As String, _
                               szTargetPath As String, _
                               szArguments As String)

    On Error GoTo Handler
    
    Dim objShell

    Dim oShellLink

    Set objShell = CreateObject("WScript.Shell")
    Set oShellLink = objShell.CreateShortcut(szShortcutPath)
    oShellLink.TargetPath = szTargetPath
    oShellLink.Arguments = szArguments
    
    oShellLink.save

    CreateShortcut = FileExists(szShortcutPath)
Handler:
End Function

Public Function GetWindowsOSVersion() As OSVERSIONINFO

    Dim osv As OSVERSIONINFO

    osv.dwOSVersionInfoSize = Len(osv)
    
    If GetVersionEx(osv) = 1 Then
        GetWindowsOSVersion = osv
    End If

End Function

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)

    Dim lState As Long

    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    
    Call SetWindowPos(frmForm.hWnd, lState, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOACTIVATE)
End Sub

Public Function StripDecimal(ByRef theNumber As Single) As Long

    Dim decimalPosition As Long

    decimalPosition = InStr(theNumber, ".")
    
    If decimalPosition > 0 Then
        StripDecimal = Left(theNumber, decimalPosition)
    Else
        StripDecimal = theNumber
    End If
    
End Function

Public Function GetFilenameFromPath(ByVal theFilePath As String) As String

    Dim theDelim As String
    
    If InStr(theFilePath, "\") > 0 Then
        theDelim = "\"
    ElseIf InStr(theFilePath, "/") > 0 Then
        theDelim = "/"
    End If
    
    GetFilenameFromPath = Right(theFilePath, Len(theFilePath) - InStrRev(theFilePath, theDelim))
End Function

Function IsPathAFolder(ByVal szPathSpec As String) As Boolean

    On Error GoTo Handler

    If Right(szPathSpec, 1) <> "\" Then
        szPathSpec = szPathSpec & "\"
    End If

    IsPathAFolder = Dir(szPathSpec) <> vbNullString
Handler:
End Function

Function AddToCollectionAtPosition(ByRef theCollection As Collection, _
                                   theItem, _
                                   desiredPosition As Long, _
                                   Optional theKey)

    If desiredPosition = 1 And theCollection.Count > 0 Then
        theCollection.Add theItem, theKey, 1
    ElseIf desiredPosition > theCollection.Count Then
        theCollection.Add theItem, theKey
    Else
        theCollection.Add theItem, theKey, , desiredPosition - 1
    End If

End Function

Function ExistInCol(ByRef cTarget As Collection, sKey) As Boolean

    On Error GoTo Handler

    ExistInCol = Not (IsEmpty(cTarget(sKey)))
    
    Exit Function

Handler:
    ExistInCol = False
End Function

Public Function RTS2(ByVal Number As Long, ByVal significance As Long)
    
    'Round number up or down to the nearest multiple of significance
    Dim d As Double
    
    Number = Number + (significance / 2)
    d = Number / significance
    d = Round(d, 0)
    RTS2 = d
End Function

'Public Function GetPngCodecCLSID() As clsid
'
'    Dim thisCLSID As New GDIPImageEncoderList
'
'    GetPngCodecCLSID = thisCLSID.EncoderForMimeType("image/png").CodecCLSID
'
'End Function

Public Function isset(srcAny) As Boolean

    On Error GoTo Handler

    Dim thisVarType As VbVarType: thisVarType = VarType(srcAny)

    If thisVarType = vbObject Then
        If Not srcAny Is Nothing Then
            isset = True

            Exit Function

        End If

    ElseIf thisVarType = vbArray Or thisVarType = 8200 Then
           
        If UBound(srcAny) > 0 Then
            isset = True

            Exit Function

        End If

    Else
        isset = IsEmpty(srcAny)

        Exit Function

    End If

Handler:
    isset = False

End Function

Public Function parseInt(srcData) As Long

    If (IsNumeric(srcData)) Then
        parseInt = CLng(srcData)

        Exit Function

    Else
        parseInt = -1
    End If

End Function

Public Function TrimNull(ByVal StrIn As String) As String

    Dim nul As Long

    ' Truncate input string at first null.
    ' If no nulls, perform ordinary Trim.
    nul = InStr(StrIn, vbNullChar)

    Select Case nul

        Case Is > 1
            TrimNull = Left(StrIn, nul - 1)

        Case 1
            TrimNull = ""

        Case 0
            TrimNull = Trim(StrIn)
    End Select

End Function

Public Function StrEnd(ByVal sData As String, _
                       ByVal sDelim As String, _
                       Optional iOffset As Integer = 1)

    If InStr(sData, sDelim) = 0 Then
        'Delim not present
    
        StrEnd = sData

        Exit Function

    End If

    Dim iLen As Integer, iDLen As Integer

    iLen = Len(sData) + 1
    iDLen = Len(sDelim)

    If iLen = 1 Or iDLen = 0 Then
        StrEnd = False

        Exit Function

    End If

    While Mid(sData, iLen, iDLen) <> sDelim And iLen > 1

        iLen = iLen - 1

    Wend

    If iLen = 0 Then
        StrEnd = False

        Exit Function

    End If
    
    StrEnd = Mid(sData, iLen + iOffset)

End Function

Public Function FileExists(sSource As String, _
                           Optional ByVal allowFsDirection As Boolean = True) As Boolean

    If sSource = vbNullString Then

        Exit Function

    End If

    Dim WFD   As WIN32_FIND_DATA

    Dim hFile As Long
    
    hFile = FindFirstFile(sSource, WFD)
    FileExists = hFile <> INVALID_HANDLE_VALUE
    
    Call FindClose(hFile)
   
    If FileExists = False And Is64bit And allowFsDirection = False Then

        Dim win64Token As Win64FSToken: Set win64Token = New Win64FSToken

        FileExists = FileExists(sSource, True)
        win64Token.EnableFS
    End If

End Function

Public Function ShowWindowTimeout(ByRef hWnd As Long, ByRef nCmdShow As ESW)

    If Not IsWindowHung(hWnd) Then
        ShowWindow hWnd, nCmdShow
    End If

End Function

Public Function IsWindowHung(hWnd As Long) As Boolean

    Dim lResult As Long

    Dim lReturn As Long
    
    lReturn = SendMessageTimeout(hWnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG Or SMTO_BLOCK, 1000, lResult)
                     
    If lReturn Then
        IsWindowHung = False

        Exit Function

    End If
    
    IsWindowHung = True

End Function

Public Function Exists(col, index) As Boolean

    On Error GoTo ExistsTryNonObject

    Dim o As Object

    Set o = col(index)
    Exists = True

    Exit Function

ExistsTryNonObject:
    Exists = ExistsNonObject(col, index)
End Function

Private Function ExistsNonObject(col, index) As Boolean

    On Error GoTo ExistsNonObjectErrorHandler

    Dim v As Variant

    v = col(index)
    ExistsNonObject = True

    Exit Function

ExistsNonObjectErrorHandler:
    ExistsNonObject = False
End Function

Public Sub RepaintWindow(ByRef hWnd As Long)

    'verified it works
    If IsWindowHung(hWnd) Then Exit Sub
    
    If hWnd <> 0 Then
        Call RedrawWindow(hWnd, ByVal 0&, ByVal 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
    End If
    
End Sub
