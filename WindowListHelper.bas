Attribute VB_Name = "MenuListHelper"
'--------------------------------------------------------------------------------
'    Component  : MenuListHelper
'    Project    : ViDock
'
'    Description: Unused code from ViGlance. Helper functions for rendering
'                 the window captions on a process
'
'--------------------------------------------------------------------------------
Option Explicit

Private Const ButtonUnpressedId As String = "unpressed"

Private Const ButtonPressedId   As String = "pressed"

Private Const ButtonNoticeId    As String = "notice"

Private Const ButtonOverId      As String = "over"

Public Enum ButtonState

    ButtonPressed = 1
    ButtonUnpressed = 2
    ButtonNotice = 3
    ButtonOver = 4

End Enum

Public Function IsMyAncestor(hWnd As Long) As Boolean

    Dim theForm As Form

    IsMyAncestor = False
    
    Set theForm = GetFormByhWnd(hWnd)

    If theForm Is Nothing Then

        Exit Function

    End If
    
    If theForm.Name = "ListMenu" Then
        IsMyAncestor = True
    End If
    
End Function

Public Function HandleWindow(hWnd As Long)

    Dim currWinP      As WINDOWPLACEMENT

    Dim MousePosition As Win.POINTL

    Dim windowPos     As Win.RECT

    If IsWindowHung(hWnd) Then Exit Function

    Call GetCursorPos(MousePosition)
    
    If GetWindowPlacement(hWnd, currWinP) > 0 Then
    
        If (currWinP.ShowCmd = SW_SHOWMINIMIZED) Then
            'minimized, so restore
            ShowWindow hWnd, SW_RESTORE
            
            g_hwndForeGroundWindow = hWnd
            SetForegroundWindow hWnd
            
            GetWindowRect hWnd, windowPos
            
            'repaint?
            MoveWindow hWnd, windowPos.Left, windowPos.Top, windowPos.Right - windowPos.Left + 1, windowPos.Bottom - windowPos.Top + 1, ByVal APITRUE
            MoveWindow hWnd, windowPos.Left, windowPos.Top, windowPos.Right - windowPos.Left, windowPos.Bottom - windowPos.Top, ByVal APITRUE
            
        ElseIf g_hwndForeGroundWindow = hWnd Then
            'normal, so minimize
            
            'AnimateWindow hwnd, 2000, AW_CENTER Or AW_HIDE
            ShowWindow hWnd, SW_MINIMIZE
        Else
        
            SetForegroundWindow hWnd
        End If
    End If

End Function

Public Function HandleListFile(ByRef theFile As ListFile)

    On Error GoTo Handler

    Shell "explorer.exe " & """" & theFile.Path & """"

    Exit Function

Handler:
End Function

Public Function GetStackContents(ByVal szStackSpec As String) As Collection

    If Right(szStackSpec, 1) <> "\" Then
        szStackSpec = szStackSpec & "\"
    End If

    Dim thisList As Collection

    Set thisList = New Collection

    Dim theFile     As ListFile

    Dim theFileName As String

    theFileName = Dir(szStackSpec)
    
    Do While theFileName > ""

        If theFileName <> vbNullString Then
            Set theFile = New ListFile
            theFile.Caption = theFileName
            theFile.Path = szStackSpec & theFileName
            
            thisList.Add theFile
        End If
        
        theFileName = Dir()
    Loop

    Set GetStackContents = thisList
End Function

Public Function DrawButtonF(ByRef theButton As Collection, _
                            ByVal theState As ButtonState, _
                            ByRef Graphics As GDIPGraphics, _
                            targetRect As gdiplus.RECTF)
    'On Error GoTo Handler:

    Dim sliceCollection As Collection

    Select Case theState
    
        Case ButtonUnpressed
            Set sliceCollection = theButton(ButtonUnpressedId)
        
        Case ButtonPressed
            Set sliceCollection = theButton(ButtonPressedId)

        Case ButtonNotice
            Set sliceCollection = theButton(ButtonNoticeId)
        
        Case ButtonOver
            Set sliceCollection = theButton(ButtonOverId)
    
    End Select
    
    If Not sliceCollection Is Nothing Then

        DrawSlicesToTargetF sliceCollection, Graphics, targetRect
    End If
    
    Exit Function

Handler:
    LogError 0, "DrawButtonF", "MenuListHelper", Err.Description
End Function

Public Function DrawButton(ByRef theButton As Collection, _
                           ByVal theState As ButtonState, _
                           ByRef Graphics As GDIPGraphics, _
                           targetRect As gdiplus.RECTL)
    'On Error GoTo Handler:

    Dim sliceCollection As Collection

    Select Case theState
    
        Case ButtonUnpressed
            Set sliceCollection = theButton(ButtonUnpressedId)
        
        Case ButtonPressed
            Set sliceCollection = theButton(ButtonPressedId)

        Case ButtonNotice
            Set sliceCollection = theButton(ButtonNoticeId)
        
        Case ButtonOver
            Set sliceCollection = theButton(ButtonOverId)
    
    End Select
    
    If Not sliceCollection Is Nothing Then

        DrawSlicesToTarget sliceCollection, Graphics, targetRect
    End If
    
Handler:
End Function

Public Function CreateButtonFromXML(ByVal szXmlElementName As String, _
                                    ByRef slicesImage As GDIPImage) As Collection
    
    Dim buttonXMlElement   As IXMLDOMElement

    Dim thisXMLNode        As IXMLDOMElement

    Dim sliceIndex         As IXMLDOMElement

    Dim buttonStateImage   As GDIPImage

    Dim state              As gdiplus.RECTL

    Dim thisButton         As Collection

    Dim buttonSlices       As Collection

    Dim szButtonIdentifier As String

    Set buttonXMlElement = ThemeHelper.GetButton(szXmlElementName)

    If buttonXMlElement Is Nothing Then
        LogError 0, "CreateButtonFromXML", "MenuListHelper", "Warning unable to load button: " & szXmlElementName

        Exit Function

    End If
    
    Set slicesImage = New GDIPImage
    slicesImage.FromFile App.Path & "\resources\" & buttonXMlElement.getAttribute("src")
    
    Set thisButton = New Collection
    
    For Each thisXMLNode In buttonXMlElement.childNodes

        If thisXMLNode.tagName = "slice_index" Then
            Set sliceIndex = thisXMLNode.cloneNode(True)
        End If

    Next
    
    If sliceIndex Is Nothing Then
        MsgBox "No slice index defined - unable to create button from XML!", vbCritical
        

        Exit Function

    End If
    
    For Each thisXMLNode In buttonXMlElement.childNodes

        If thisXMLNode.tagName = "state" Then
        
            If Not IsNull(thisXMLNode.getAttribute("id")) Then
                szButtonIdentifier = CStr(thisXMLNode.getAttribute("id"))
            End If
        
            state.Left = CLng(thisXMLNode.getAttribute("x"))
            state.Top = CLng(thisXMLNode.getAttribute("y"))
            
            state.Width = CLng(thisXMLNode.getAttribute("width"))
            state.Height = CLng(thisXMLNode.getAttribute("height"))
            
            Set buttonStateImage = CreateNewImageFromSection(slicesImage, state)
            Set buttonSlices = CreateSlicesFromXMLElement(sliceIndex, buttonStateImage)
            
            If szButtonIdentifier <> vbNullString Then
                thisButton.Add buttonSlices, szButtonIdentifier
            End If
        End If

    Next
    
    Set CreateButtonFromXML = thisButton
End Function

Public Function FindMaxHeight(ByRef ListMenu As Collection, ByVal itemDifference As Long)

    On Error GoTo Handler

    FindMaxHeight = ListMenu.Count * itemDifference

Handler:

    Exit Function

End Function

Public Function FindMaxWidth(ByRef ListMenu As Collection, _
                             ByRef theGraphics As GDIPGraphics, _
                             ByRef theFont As GDIPFont)

    On Error GoTo Handler

    Dim thisItem As Object

    Dim lpRect   As gdiplus.RECTF

    Dim maxWidth As Long

    For Each thisItem In ListMenu

        lpRect = theGraphics.MeasureString(thisItem.Caption, theFont)
        Debug.Print lpRect.Width
        
        If lpRect.Width > maxWidth Then
            maxWidth = lpRect.Width
        End If

    Next
    
    FindMaxWidth = maxWidth
    
Handler:

    Exit Function

End Function
