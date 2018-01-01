Attribute VB_Name = "AppHelper"
'--------------------------------------------------------------------------------
'    Component  : AppHelper
'    Project    : ViDock
'
'    Description: Functions that had nowhere else to live :(
'                 TODO: Create dedicated modules for each function no matter
'                 how slim they might be
'
'--------------------------------------------------------------------------------

Option Explicit

Public Const WINDOW_IMAGE_HEIGHT As Long = 120

Public Const WINDOW_IMAGE_WIDTH  As Long = 187

Public Const JUMPLIST_CAP        As Long = 9

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
    MAKELPARAM = MakeLong(wLow, wHigh)
End Function

Public Function MAKEWPARAM(LOWWORD As Long, HiWord As Long) As Long
    MAKEWPARAM = (LOWWORD And &HFFFF&) Or (HiWord * &H10000)
End Function

'Changes MousePosition if it's inside Object
Public Function MouseInsideObject(ByRef MousePosition As POINTS, _
                                  Left As Long, _
                                  Top As Long, _
                                  Width As Long, _
                                  Height As Long) As Boolean

    If MousePosition.X > Left And MousePosition.X < Left + Width And MousePosition.Y > Top And MousePosition.Y < Top + Height Then
        
        MouseInsideObject = True
        MousePosition.X = MousePosition.X - Left
        MousePosition.Y = MousePosition.Y - Top
    End If
    
End Function

Public Function ExitApplication()

    Dim thisForm As Form

    For Each thisForm In Forms

        Unload thisForm
    Next

End Function

Public Function DebugInfo_IsSet(ByRef theObject, _
                                ByVal theObjectName As String, _
                                ByRef theText As String)

    If Not theObject Is Nothing Then
        theText = theText & theObjectName & " = Initialzed"
    Else
        theText = theText & theObjectName & " = Null"
    End If

    theText = theText & vbCrLf
End Function

Public Function GetFormByhWnd(ByVal hWnd As Long) As Form

    Dim thisForm As Form

    For Each thisForm In Forms

        If thisForm.hWnd = hWnd Then
            Set GetFormByhWnd = thisForm

            Exit For

        End If

    Next
    
End Function

Public Function hWndBelongToUs(hWnd As Long, Optional ExceptionHwnd As Long) As Boolean

    Dim thisForm As Form

    hWndBelongToUs = False

    For Each thisForm In Forms

        If thisForm.hWnd = hWnd Then
            If hWnd = ExceptionHwnd Then
                hWndBelongToUs = False
            Else
                hWndBelongToUs = True
            End If
            
            Exit For

        End If

    Next
    
End Function

Function RunningInVB() As Boolean

    'Returns whether we are running in vb(true), or compiled (false)
    Static counter As Variant

    If IsEmpty(counter) Then
        counter = 1
        Debug.Assert RunningInVB() Or True
        counter = counter - 1
    ElseIf counter = 1 Then
        counter = 0
    End If

    RunningInVB = counter
End Function

Public Function ShowStartMenu()
    SendMessage FindWindow("Shell_TrayWnd", ""), ByVal WM_SYSCOMMAND, ByVal SC_TASKLIST, ByVal 0
End Function

Public Function SetOwner(ByRef ownerhWnd As Long, windowToOwnhWnd As Long) As Long
    SetOwner = SetWindowLong(windowToOwnhWnd, GWL_HWNDPARENT, ownerhWnd)
End Function

Public Function TrackMouseEvents(hWnd As Long) As Boolean

    Dim ET As TrackMouseEvent

    TrackMouseEvents = False
    
    'initialize structure
    ET.cbSize = Len(ET)
    ET.hwndTrack = hWnd
    ET.dwFlags = TME_LEAVE

    'start the tracking
    If Not TrackMouseEvent(ET) = 0 Then
        TrackMouseEvents = True
    End If
    
End Function

Public Function ShellExec(szProgram As String, Optional szParams As String)
    ShellExecute 0, "Open", szProgram, szParams, App.Path, 1
End Function

Public Function ExtractXMLTextElement(ByRef parentElement As IXMLDOMElement, _
                                      ByVal szElementName As String, _
                                      ByVal DefaultValue As String) As String

    On Error GoTo Handler
    
    ExtractXMLTextElement = CStr(parentElement.selectSingleNode(szElementName).Text)

    Exit Function

Handler:
    ExtractXMLTextElement = DefaultValue
End Function

Public Function CreateXMLTextElement(ByRef sourceDoc As DOMDocument, _
                                     ByRef parentElement As IXMLDOMElement, _
                                     ByVal szElementName As String, _
                                     ByVal szValue As String)
    
    Dim element As IXMLDOMElement

    Set element = sourceDoc.createElement(szElementName)
    parentElement.appendChild element
    
    element.Text = szValue
End Function

Public Sub LoadWindowsPinnedApps(ByRef theTaskBar As TaskBar)
    
    Dim theFileName As String

    Dim basePath    As String

    basePath = Environ("appdata") & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\"
    theFileName = Dir(basePath)

    Do While theFileName <> vbNullString
        theTaskBar.AddPinnedApp basePath & theFileName
        theFileName = Dir()
    Loop

End Sub

Public Function ReadPinnedPrograms(ByRef programsXML As IXMLDOMElement, _
                                   ByRef pinnedProcesses As Collection)

    On Error GoTo Handler

    Dim thisXMLProgram As IXMLDOMElement

    Dim thisProgram    As Process

    Dim szIconPath     As String

    'programsXML.Load sCon_AppDataPath & "programs.xml"

    For Each thisXMLProgram In programsXML.childNodes

        If Not IsNull(thisXMLProgram.getAttribute("path")) Then
            Set thisProgram = New Process
            szIconPath = vbNullString
            
            If thisXMLProgram.tagName = "stack" Then
                thisProgram.IsStack = True
                
                thisProgram.Constructor 0, Environ("windir") & "\explorer.exe"
                thisProgram.CreateIconFromPath
                
                thisProgram.Path = thisXMLProgram.getAttribute("path")
            Else
                thisProgram.Constructor 0, thisXMLProgram.getAttribute("path")
                thisProgram.CreateIconFromPath
            End If
            
            If Not IsNull(thisXMLProgram.getAttribute("arguments")) Then
                thisProgram.Arguments = thisXMLProgram.getAttribute("arguments")
            End If
        
            If Not IsNull(thisXMLProgram.getAttribute("caption")) Then
                thisProgram.Caption = thisXMLProgram.getAttribute("caption")
            End If
            
            thisProgram.Pinned = True
            
            pinnedProcesses.Add thisProgram
        End If

    Next

    Exit Function

Handler:
End Function

Public Function DumpPinnedProcesses(ByRef sourceDoc As DOMDocument, _
                                    ByRef parentElement As IXMLDOMElement, _
                                    ByRef pinnedProcesses As Collection) As Boolean

    Dim XML_programs As IXMLDOMElement

    Dim thisProgram  As Process

    Dim newItem      As IXMLDOMElement

    'Set m_sourceDoc = New DOMDocument
    
    Set XML_programs = sourceDoc.createElement("pinned_programs")
    parentElement.appendChild XML_programs
    
    For Each thisProgram In pinnedProcesses

        If thisProgram.IsStack Then
            Set newItem = sourceDoc.createElement("stack")
        Else
            Set newItem = sourceDoc.createElement("program")
        End If
        
        XML_programs.appendChild newItem
        
        newItem.setAttribute "path", thisProgram.Path
        newItem.setAttribute "caption", thisProgram.Caption
        
        If Not thisProgram.IsStack Then newItem.setAttribute "arguments", thisProgram.Arguments
    Next
    
    DumpPinnedProcesses = True
End Function

Public Function GetBestProcessCaption(ByVal theProcess As Process) As String

    Dim thisWindow As Window

    If theProcess.WindowCount = 1 Then
        Set thisWindow = theProcess.Window(1)
        GetBestProcessCaption = thisWindow.Caption
    ElseIf theProcess.Caption <> vbNullString Then
        GetBestProcessCaption = theProcess.Caption
    Else
        GetBestProcessCaption = theProcess.ImageName
    End If

End Function

