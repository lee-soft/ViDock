VERSION 5.00
Begin VB.Form TaskBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   14
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timUpdateClock 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer timEnumerateTasks 
      Interval        =   400
      Left            =   4560
      Top             =   1200
   End
End
Attribute VB_Name = "TaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : TaskBar
'    Project    : ViDock
'
'    Description: The physical dock object
'
'--------------------------------------------------------------------------------
Option Explicit

Implements IHookSink

Private Const EFFECT_MARGIN          As Single = 3

Private Const ACTUAL_ICON_SIZE       As Long = 16

Private Const SEPERATOR_SIZE         As Long = 5

Private Const DOCK_MARGIN_X          As Long = 100

Private m_dockPopup                  As DockPopup

Private WithEvents m_listMenu        As ListMenu
Attribute m_listMenu.VB_VarHelpID = -1

Private WithEvents m_listMenuPointer As ListMenuPointer
Attribute m_listMenuPointer.VB_VarHelpID = -1

Private m_TaskbarBackgroundGraphics  As GDIPGraphics

Private m_TaskbandGraphics           As GDIPGraphics

Private m_TaskbandBitmap             As GDIPBitmap

Private m_backgroundImage            As GDIPImage

Private m_separator                  As GDIPImage

Private m_indicatorImage             As GDIPImage

Private m_indicatorSlices            As Collection

Private WithEvents m_menuBar         As NewMenuBar
Attribute m_menuBar.VB_VarHelpID = -1

Private m_systemTrayPopup            As TrayPopup

Private WithEvents m_clock           As SystemClock
Attribute m_clock.VB_VarHelpID = -1

Private m_arrowImage                 As GDIPImage

Private m_arrowSlices                As Collection

Private m_arrowPlacement             As gdiplus.RECTL

Private m_arrowState                 As ButtonState

Private m_taskList                   As TaskList

Private m_currentItemCount           As Long

Private m_processIndex               As Long

Private m_layeredWindowProperties    As LayerdWindowHandles

Private m_taskbandDimensions         As gdiplus.RECTL

Private m_visibleList                As Collection

Private m_sizeFactor                 As Double

Private m_lastSelectedProcess        As Process

Private m_lastProcessShowcased       As Process

Private m_objectBeingMouseTracked    As Object

Private m_initialized                As Boolean

Private m_Shell_Hook_Msg_ID          As Long

Private m_processArea                As Object

Private m_groupMenu                  As clsMenu

Private WithEvents m_startButton As StartOrb
Attribute m_startButton.VB_VarHelpID = -1

Private m_slices                 As Collection

Private m_menuBarSlice           As Slice

Private m_mouseTracking          As Boolean

Private m_ignoreTrayMenuSpawn    As Boolean

Private m_componentOrder         As Collection

Private m_mousePosition          As POINTS

Sub ApplyGlass()

    Dim a  As Margins '??????????????

    Dim bb As DWM_BLURBEHIND

    Dim hr As Long

    'make entire region glass (for invisible effect)
    a.m_Left = -1
    a.m_Top = -1
    a.m_Right = -1
    a.m_Bottom = -1
    
    DwmIsCompositionEnabled 1
    
    Me.BackColor = vbBlack
    'If DwmExtendFrameIntoClientArea(Me.hWnd, a) <> S_OK Then
    'MsgBox "Error!"
    'End If

End Sub

Private Sub Form_Activate()

    If m_systemTrayPopup.Visible Then
        m_systemTrayPopup.Hide
    End If

End Sub

Sub AddPinnedApplicationsFirst()

    On Error GoTo Handler
    
    Dim thisProcess As Process

    For Each thisProcess In Settings.PinnedApplications

        m_taskList.Processes.Add thisProcess, thisProcess.GetKey
    Next
    
Handler:
End Sub

Private Sub Form_DragDropStack(szStackSpec As String)
    'On Error GoTo Handler
    
    Dim stackCaption As String

    stackCaption = StrEnd(szStackSpec, "\")

    If IsPathAFolder(szStackSpec) Then

        Dim newProcess As New Process

        newProcess.Constructor 0, Environ("windir") & "\explorer.exe"
        newProcess.CreateIconFromPath
        newProcess.Path = szStackSpec
        newProcess.Caption = stackCaption
        newProcess.IsStack = True

        If Not Exists(m_taskList.Processes, newProcess.GetKey) Then
            newProcess.Pinned = True
        
            AddToCollectionAtPosition m_taskList.Processes, newProcess, 1, newProcess.GetKey
            AddToCollectionAtPosition Settings.PinnedApplications, newProcess, 1
        End If
    End If

    Exit Sub

Handler:
    MsgBox Err.Description
End Sub

Sub Form_DragDropFile(sFile As String)

    On Error GoTo Handler

    If Not AddPinnedApp(sFile) Then Exit Sub
    
    Dim szFileName As String

    szFileName = StrEnd(sFile, "\")
    
    FileCopy sFile, Environ("appdata") & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\" & szFileName

    Exit Sub

Handler:
    LogError 0, "DragDropFile", "TaskBar", Err.Description
End Sub

Function AddPinnedApp(sFile As String) As Boolean

    On Error GoTo Handler
    
    Dim thisLink      As ShellLink

    Dim backupCaption As String

    Dim extension     As String

    thisLink = GetShortcut(CStr(sFile))
    
    If InStr(sFile, ".") > 0 And InStr(sFile, "\") Then
        backupCaption = StrEnd(sFile, "\")
        extension = StrEnd(backupCaption, ".")
        backupCaption = Mid(backupCaption, 1, Len(backupCaption) - Len(extension) - 1)
    End If

    If FileExists(thisLink.szPath) Then

        Dim newProcess As New Process

        newProcess.Constructor 0, thisLink.szPath
        newProcess.CreateIconFromPath
        newProcess.Arguments = thisLink.szArguments
        newProcess.PhysicalLinkFile = sFile
        
        If newProcess.Caption = vbNullString Then
            newProcess.Caption = backupCaption
        End If
        
        If Not Exists(m_taskList.Processes, newProcess.GetKey) Then
            newProcess.Pinned = True
        
            AddToCollectionAtPosition m_taskList.Processes, newProcess, 1, newProcess.GetKey
            AddToCollectionAtPosition Settings.PinnedApplications, newProcess, 1
        End If
    End If

    AddPinnedApp = True

    Exit Function

Handler:
    
    MsgBox Err.Description

End Function

Sub Initialize()

    If Me.Left <> 0 Or Me.Top <> 0 Or Me.Width <> Screen.Width Or Me.Height <> (36 * Screen.TwipsPerPixelY) Then
        Me.Move 0, 0, Screen.Width, 36 * Screen.TwipsPerPixelY
    End If

    Set m_startButton = New StartOrb
    Set m_taskList = New TaskList
    Set m_dockPopup = New DockPopup
    Set m_listMenu = New ListMenu
    Set m_listMenuPointer = New ListMenuPointer
    Set m_systemTrayPopup = New TrayPopup
    Set m_clock = New SystemClock
    Set m_processArea = Me
    Set m_componentOrder = New Collection
    
    Set m_arrowSlices = MenuListHelper.CreateButtonFromXML("tray_arrow", m_arrowImage)
    Set m_indicatorSlices = MenuListHelper.CreateButtonFromXML("indicator", m_indicatorImage)
    
    If m_indicatorSlices.Count > 0 Then
        'Set m_indicatorFirstSlice = m_indicatorSlices(1)
    End If
    
    m_arrowPlacement = CreateRectL(16, 16, Me.ScaleWidth - 19, 5)
    m_arrowState = ButtonUnpressed
    
    Set m_menuBar = New NewMenuBar
    
    Set m_backgroundImage = New GDIPImage
    Set m_separator = New GDIPImage
    'Set m_indicator = New GDIPImage
    'Set m_textPath = New GDIPGraphicPath
    
    Set m_visibleList = New Collection
    m_separator.FromFile App.Path & "\resources\separator.png"
    'm_indicator.FromFile App.Path & "\resources\indicator.png"
    
    Set m_slices = SliceHelper.CreateSlicesFromXML("background", m_backgroundImage)

    If ExistInCol(m_slices, "bar") Then
        Set m_menuBarSlice = m_slices("bar")
    End If
    
    Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    ' ShowWindow TaskbarHandler, SW_HIDE
   
    'Set m_systemTray = New SystemTrayManager
    'Set m_systemTray.HostForm = Me
    'm_systemTray.Popup = m_dockPopup
    'm_systemTray.Dimensions = CreateRectL(0, 0, 0, 3)

    Set m_menuBar.HostForm = Me
    m_menuBar.Dimensions = CreateRectL(Me.ScaleHeight, Me.ScaleWidth - 300, 50, 0)
    'm_menuBar.Dimensions = CreateRectF(50, 0, Me.ScaleHeight, Me.ScaleWidth - 300)
    m_menuBar.Background = m_backgroundImage

    'Set owner of the window to me
    SetOwner Me.hWnd, m_startButton.hWnd
    
    m_startButton.Show
    
    If Not m_menuBarSlice Is Nothing Then
        m_startButton.Move 0, (m_menuBarSlice.Height / 2 - m_startButton.ScaleHeight / 2) * Screen.TwipsPerPixelY
    End If

    AppHelper.LoadWindowsPinnedApps Me
    AddPinnedApplicationsFirst
    
    m_initialized = True

End Sub

Private Sub ResizeTaskBandIfNeeded(newItemNumber As Long)

    If m_currentItemCount = newItemNumber Then

        Exit Sub

    End If

    Set m_TaskbandGraphics = New GDIPGraphics
    Set m_TaskbandBitmap = New GDIPBitmap

    m_TaskbandBitmap.CreateFromSizeFormat newItemNumber * (ACTUAL_ICON_SIZE + SEPERATOR_SIZE), (ACTUAL_ICON_SIZE + SEPERATOR_SIZE), GDIPlusWrapper.Format32bppArgb
    m_TaskbandGraphics.FromImage m_TaskbandBitmap.Image

    m_currentItemCount = newItemNumber
    
    CalculateTaskbandDimensions
    ResetProcessPositions
End Sub

Private Sub Form_Initialize()

    If GetWindowsOSVersion.dwMajorVersion > 5 Then
        ChangeWindowMessageFilter WM_DROPFILES, MSGFLT_ADD
        ChangeWindowMessageFilter WM_COPYDATA, MSGFLT_ADD
        ChangeWindowMessageFilter &H49, MSGFLT_ADD
    End If

    DragAcceptFiles Me.hWnd, APITRUE
End Sub

Private Sub Form_Load()
    Initialize
    
    'ApplyDWMEffect
    CalculateTaskbandDimensions

    HookWindow Me.hWnd, Me
    
    RegisterShellHookWindow Me.hWnd
    m_Shell_Hook_Msg_ID = WinAPIHelper.RegisterWindowMessage("SHELLHOOK")
    
    TrackMouseEvents Me.hWnd
    'StayOnTop Me, True
End Sub

Private Function ReleaseGraphics()
    m_TaskbarBackgroundGraphics.ReleaseHDC m_layeredWindowProperties.theDC
    'm_TaskbandGraphics.ReleaseHDC
End Function

Private Function InitializeGraphics()
    Set m_TaskbarBackgroundGraphics = New GDIPGraphics
    
    'm_graphics.FromHDC Me.hdc
    m_TaskbarBackgroundGraphics.FromHDC m_layeredWindowProperties.theDC
    
    'm_TaskbarBackgroundGraphics.SmoothingMode = SmoothingModeHighQuality
    'm_TaskbarBackgroundGraphics.InterpolationMode = InterpolationModeHighQualityBicubic
    
    m_TaskbarBackgroundGraphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    m_TaskbarBackgroundGraphics.SmoothingMode = SmoothingModeHighQuality
    m_TaskbarBackgroundGraphics.PixelOffsetMode = PixelOffsetModeHighQuality
    'm_TaskbarBackgroundGraphics.CompositingMode = CompositingModeSourceCopy
    
    'm_TaskbarBackgroundGraphics.CompositingMode = CompositingModeSourceOver
    'm_TaskbarBackgroundGraphics.CompositingQuality = CompositingQualityHighQuality
    m_TaskbarBackgroundGraphics.InterpolationMode = InterpolationModeNearestNeighbor

    m_clock.Initialize m_TaskbarBackgroundGraphics
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim MousePosition As POINTS

    MousePosition.X = X
    MousePosition.Y = Y
    
    If MouseInsideObject(MousePosition, m_menuBar.Dimensions.Left, m_menuBar.Dimensions.Top, m_menuBar.Dimensions.Width, m_menuBar.Dimensions.Height) Then
        m_menuBar.MouseDown Button, CSng(MousePosition.X), CSng(MousePosition.Y)
    End If
    
    If m_ignoreTrayMenuSpawn Then
        If Not MouseInsideObject(MousePosition, m_arrowPlacement.Left, m_arrowPlacement.Top, m_arrowPlacement.Width, m_arrowPlacement.Height) Then
            m_ignoreTrayMenuSpawn = False
        End If
    End If
    
End Sub

Private Sub ResetObjectBeingTracked()

    If m_objectBeingMouseTracked Is Nothing Then Exit Sub

    If m_objectBeingMouseTracked Is m_clock Then
        m_clock.MouseLeft
    ElseIf m_objectBeingMouseTracked Is m_menuBar Then
        m_menuBar.MouseLeft
    End If

    Set m_objectBeingMouseTracked = Nothing
End Sub

Private Sub SetObjectBeingTracked(ByRef newObject As Object)
    
    If Not m_objectBeingMouseTracked Is Nothing Then
        If newObject Is m_objectBeingMouseTracked Then

            Exit Sub

        End If
        
        If m_objectBeingMouseTracked Is m_clock Then
            m_clock.MouseLeft
        ElseIf m_objectBeingMouseTracked Is m_menuBar Then
            m_menuBar.MouseLeft
        ElseIf m_objectBeingMouseTracked Is m_processArea Then
            Debug.Print "HandleMouseLeavingProcessArea"

            HandleMouseLeavingProcessArea
        End If

    End If
    
    Debug.Print "Changing tracking to: " & TypeName(newObject)
    Set m_objectBeingMouseTracked = newObject
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_mouseTracking = False Then
        m_mouseTracking = TrackMouseEvents(Me.hWnd)
        'HandleMouseEnter
    End If

    m_mousePosition.X = X
    m_mousePosition.Y = Y

    If MouseInsideObject(m_mousePosition, m_clock.X, 0, m_clock.Width, Me.ScaleHeight) Then
        SetObjectBeingTracked m_clock
        
        m_clock.MouseMove Button, X, Y
        
    ElseIf MouseInsideObject(m_mousePosition, m_arrowPlacement.Left, m_arrowPlacement.Top, m_arrowPlacement.Width, m_arrowPlacement.Height) Then
        'ResetObjectBeingTracked
        SetObjectBeingTracked timUpdateClock
        
        If m_arrowState <> ButtonPressed Then
            m_arrowState = ButtonPressed
            DrawTaskList
        End If

    ElseIf Not GetSelectedProcess(X) Is Nothing Then
        SetObjectBeingTracked m_processArea
    
        HandleProcessSelection X
    
    ElseIf MouseInsideObject(m_mousePosition, m_menuBar.Dimensions.Left, m_menuBar.Dimensions.Top, m_menuBar.Dimensions.Width, m_menuBar.Dimensions.Height) Then
        SetObjectBeingTracked m_menuBar
        
        m_menuBar.MouseMove m_mousePosition
    Else

        If Not m_systemTrayPopup.Visible Then
            If m_arrowState <> ButtonUnpressed Then
                m_arrowState = ButtonUnpressed
                DrawTaskList
            End If
        End If
    End If

End Sub

Private Sub HandleMouseLeavingProcessArea()
    Set m_lastSelectedProcess = Nothing
    m_dockPopup.Hide
End Sub

Private Sub HandleProcessSelection(ByVal X As Single)

    Dim windowListMenuOpen As Boolean

    Dim selectedProcess    As Process

    If Not m_listMenuPointer Is Nothing Then
        If m_listMenuPointer.Visible Then
            windowListMenuOpen = True
        End If
    End If

    Set selectedProcess = GetSelectedProcess(X)
    
    If m_lastSelectedProcess Is selectedProcess Then

        Exit Sub

    End If
    
    Set m_lastSelectedProcess = selectedProcess

    If selectedProcess Is Nothing Then
        Set m_lastProcessShowcased = Nothing

        Exit Sub

    End If
    
    If Not m_lastProcessShowcased Is Nothing Then
        If m_lastProcessShowcased Is selectedProcess Then

            Exit Sub

        End If
    End If
    
    If Not windowListMenuOpen And Not m_systemTrayPopup.Visible Then
        m_dockPopup.Top = Me.Top + Me.Height
        
        m_dockPopup.ShowTextPopup GetBestProcessCaption(selectedProcess), m_taskbandDimensions.Left + (Me.Left / Screen.TwipsPerPixelX) + (selectedProcess.X) + (ACTUAL_ICON_SIZE / 2)
                                                            
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Dim isInsideTaskband As Boolean
    Dim MousePosition As POINTS

    MousePosition.X = X
    MousePosition.Y = Y
    
    If MouseInsideObject(MousePosition, m_arrowPlacement.Left, m_arrowPlacement.Top, m_arrowPlacement.Width, m_arrowPlacement.Height) Then
        If m_ignoreTrayMenuSpawn Then
            m_ignoreTrayMenuSpawn = False

            Exit Sub

        End If
        
        m_systemTrayPopup.Top = Me.Top + Me.Height
        m_systemTrayPopup.ShowTrayPopup m_arrowPlacement.Left

        Exit Sub

    End If
    
    If HandleTaskbarMouseUp(Button) Then Exit Sub

    If Button = vbRightButton Then
        SystemMenuHandler SystemMenu.ShowMenu(Me.hWnd)

        Exit Sub

    End If

End Sub

Private Function HandleTaskbarMouseUp(Button As Integer) As Boolean

    Dim singleWindow As Window

    If Button = vbLeftButton Then
    
        If m_lastSelectedProcess Is Nothing Then Exit Function
        HandleTaskbarMouseUp = True
    
        If m_lastSelectedProcess.IsStack Then
        
            m_dockPopup.Hide
        
            If (Not m_lastProcessShowcased Is Nothing) Or m_lastProcessShowcased Is m_lastSelectedProcess Then
                m_listMenuPointer.Hide
                Set m_lastProcessShowcased = Nothing
                Set m_lastSelectedProcess = Nothing
            Else
                Set m_lastProcessShowcased = m_lastSelectedProcess
                m_listMenuPointer.ShowList GetStackContents(m_lastSelectedProcess.Path), Me.Top, m_taskbandDimensions.Left + (Me.Left / Screen.TwipsPerPixelX) + (m_lastSelectedProcess.X * m_sizeFactor) + ((ACTUAL_ICON_SIZE * m_sizeFactor) / 2)
            End If
        
        ElseIf Not m_lastSelectedProcess.Running Then
     
            ShellExec m_lastSelectedProcess.Path, m_lastSelectedProcess.Arguments
        
        Else
    
            If m_lastSelectedProcess.WindowCount = 1 Then
                
                Set singleWindow = m_lastSelectedProcess.Window(1)
                HandleWindow singleWindow.hWnd
                
            Else
            
                m_dockPopup.Hide
                
                If m_lastProcessShowcased Is m_lastSelectedProcess Then
                    Set m_lastProcessShowcased = Nothing
                    m_listMenuPointer.Hide
                Else
                    Set m_lastProcessShowcased = m_lastSelectedProcess
                    m_listMenuPointer.ShowList m_lastSelectedProcess.Window, Me.Top + Me.Height, m_taskbandDimensions.Left + (Me.Left / Screen.TwipsPerPixelX) + (m_lastSelectedProcess.X) + (ACTUAL_ICON_SIZE / 2), vbPopupMenuLeftAlign
                End If
            End If
        
        End If
        
    ElseIf Button = vbRightButton Then

        'Set m_currentJumpLists = m_lastSelectedProcess.GetJumpLists
        If Not m_lastSelectedProcess Is Nothing Then
        
            Set m_groupMenu = BuildMenuWithoutJumpList()
            m_groupMenu.EditItem 5, IIf(m_lastSelectedProcess.Pinned, "Unpin", "Pin")
            HandleProcessMenuResult m_groupMenu.ShowMenu(Me.hWnd), m_lastSelectedProcess
            
            HandleTaskbarMouseUp = True
        End If
        
    End If
    
End Function

Private Sub Form_OLEDragDrop(data As DataObject, _
                             Effect As Long, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Dim j As Long

    If data.GetFormat(vbCFFiles) = True Then

        For j = 1 To data.Files.Count

            If (GetAttr(data.Files.Item(j)) And vbDirectory) Then
                Form_DragDropStack data.Files.Item(j)
            Else
                Form_DragDropFile data.Files.Item(j)
            End If

        Next

    End If
    
End Sub

Private Sub Form_OLEDragOver(data As DataObject, _
                             Effect As Long, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single, _
                             state As Integer)

    Dim theFileExtension As String

    Debug.Print data.Files.Item(1)

    Effect = vbDropEffectNone

    If Not data.GetFormat(vbCFFiles) = True Then

        Exit Sub

    End If

    If Not (GetAttr(data.Files.Item(1)) = vbDirectory) Then
        theFileExtension = UCase(StrEnd(data.Files.Item(1), "."))
    
        If Not theFileExtension = "LNK" Then

            Exit Sub

        End If
    End If
    
    Effect = vbDropEffectCopy
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_menuBar = Nothing
    Set m_taskList = Nothing
    
    Unload m_dockPopup
    Unload m_listMenu
    Unload m_startButton
    
    ReleaseGraphics
    UnhookWindow Me.hWnd
    
    Settings.Comit
    
    ReleaseGraphics
    m_layeredWindowProperties.Release
    
    ExitApplication
End Sub

Private Sub Form_Resize()

    If Not m_initialized Then Exit Sub

    InitializeGraphics
    CalculateTaskbandDimensions
End Sub

Sub CalculateTaskbandDimensions()

    If m_TaskbandBitmap Is Nothing Then Exit Sub
    
    'm_systemTray.Dimensions = CreateRectL(0, 0, Me.ScaleWidth - m_systemTray.Image.Width, 3)

    m_taskbandDimensions.Width = m_TaskbandBitmap.Image.Width
    m_taskbandDimensions.Height = m_TaskbandBitmap.Image.Height
    
    m_taskbandDimensions.Left = Me.ScaleWidth - ACTUAL_ICON_SIZE - SEPERATOR_SIZE - m_taskbandDimensions.Width - m_clock.Width - 3
    m_clock.X = Me.ScaleWidth - (m_clock.Width + SEPERATOR_SIZE)
    m_clock.Y = 2
    
    m_arrowPlacement.Left = m_clock.X - (ACTUAL_ICON_SIZE + SEPERATOR_SIZE)
    
    m_menuBar.Dimensions = CreateRectL(Me.ScaleHeight, m_taskbandDimensions.Left - 50, 50, 0)
    m_menuBar.Background = m_backgroundImage

    m_taskbandDimensions.Top = 0

End Sub

Private Function IHookSink_WindowProc(hWnd As Long, _
                                      msg As Long, _
                                      wp As Long, _
                                      lp As Long) As Long
    
    'Dim bConsume As Boolean
    Dim theWindow As Window

    On Error GoTo Handler
    
    If msg = WM_MOUSELEAVE Then
        m_mouseTracking = False
        
        HandleMouseLeave
        
    ElseIf msg = m_Shell_Hook_Msg_ID Then

        'wp = 32772 (for windows 8)
        If wp = HSHELL_WINDOWACTIVATED Or wp = 32772 Then
            
            HandleWindowActivated lp
            
        ElseIf wp = HSHELL_FLASH Then
        
            Set theWindow = m_taskList.GetWindowByHWND(lp)

            If Not theWindow Is Nothing Then
                theWindow.Flashing = True
            End If
        End If

    ElseIf msg = WM_ACTIVATE Then

        If wp = WA_ACTIVE Or wp = WA_CLICKACTIVE Then
            
            If lp = m_systemTrayPopup.hWnd Then
                m_ignoreTrayMenuSpawn = True
            End If
            
        End If
    
    ElseIf msg = WM_DROPFILES Then

        Dim hFilesInfo  As Long

        Dim szFileName  As String

        Dim wTotalFiles As Long

        Dim wIndex      As Long
        
        hFilesInfo = wp
        wTotalFiles = DragQueryFileW(hFilesInfo, &HFFFF, ByVal 0&, 0)
    
        For wIndex = 0 To wTotalFiles
            szFileName = Space$(1024)
            
            If Not DragQueryFileW(hFilesInfo, wIndex, StrPtr(szFileName), Len(szFileName)) = 0 Then
                Form_DragDropFile TrimNull(szFileName)
            End If

        Next wIndex
        
        DragFinish hFilesInfo
    End If

    'If bConsume Then Exit Function
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = InvokeWindowProc(hWnd, msg, wp, lp)
End Function

Private Function HandleMouseLeave()
    m_menuBar.ResetSelector
End Function

Private Sub HandleWindowActivated(hWnd As Long)

    If hWnd = 0 Then

        Exit Sub

    End If

    If hWndBelongToUs(hWnd) Then

        Exit Sub

    End If
    
    m_menuBar.PopulateFromhWnd hWnd

End Sub

Private Sub m_clock_onMouseLeaves()
    m_dockPopup.Hide
End Sub

Private Sub m_clock_onPopup(ByVal szText As String)

    If Not m_dockPopup.Visible Then
        m_dockPopup.Top = Me.Top + Me.Height
        m_dockPopup.ShowTextPopup szText, m_clock.X + (m_clock.Width / 2)
    End If

End Sub

Private Sub m_listMenu_onClosed()
    'Set m_lastProcessShowcased = Nothing
    Set m_lastSelectedProcess = Nothing
End Sub

Private Sub m_listMenuPointer_onRightClick(theItem As Object)

    Dim targetWindow          As Window

    Dim theMenuHandle         As Long

    Dim thisMenu              As clsMenu

    Dim sysCmdID              As Long

    Dim currentCursorPosition As Win.POINTL

    Set targetWindow = theItem

    If targetWindow.IsHung Then Exit Sub

    GetCursorPos currentCursorPosition

    theMenuHandle = GetSystemMenu(targetWindow.hWnd, 0)
    Set thisMenu = CreateSystemMenu(theMenuHandle, targetWindow.WindowState)

    SetForegroundWindow m_listMenuPointer.hWnd
    sysCmdID = thisMenu.ShowMenu(m_listMenuPointer.hWnd)

    Select Case sysCmdID
    
        Case SC_RESTORE
            ShowWindow targetWindow.hWnd, SW_SHOWNORMAL
    
        Case SC_MINIMIZE
            ShowWindow targetWindow.hWnd, SW_SHOWMINIMIZED
    
        Case SC_MAXIMIZE
            ShowWindow targetWindow.hWnd, SW_SHOWMAXIMIZED
        
        Case SC_CLOSE
            PostMessage targetWindow.hWnd, WM_CLOSE, 0&, 0&
        
        Case Else
            PostMessage targetWindow.hWnd, WM_SYSCOMMAND, ByVal sysCmdID, ByVal MAKELPARAM(currentCursorPosition.X, currentCursorPosition.Y)
    
    End Select

End Sub

Private Sub m_menuBar_onChanged()
    DrawTaskList
    
End Sub

Private Sub m_startButton_onMove(newX As Long, newY As Long)
    newX = 0
    newY = (m_menuBarSlice.Height / 2 - m_startButton.ScaleHeight / 2)
End Sub

Private Sub m_systemTray_onChange()
    CalculateTaskbandDimensions
End Sub

Private Sub timEnumerateTasks_Timer()

    Dim cursorPos As POINTL

    GetCursorPos cursorPos
    
    If Me.Left <> 0 Or Me.Top <> 0 Or Me.Width <> Screen.Width Or Me.ScaleHeight <> m_backgroundImage.Height Then
        Me.Move 0, 0, Screen.Width, m_backgroundImage.Height * Screen.TwipsPerPixelY
    End If
    
    If cursorPos.X > (Me.Left + Me.Width) / (Screen.TwipsPerPixelX) Or cursorPos.X < (Me.Left / Screen.TwipsPerPixelX) Or cursorPos.Y > (Me.Top + Me.Height) / (Screen.TwipsPerPixelY) Or cursorPos.Y < (Me.Top / Screen.TwipsPerPixelY) Then
        
        m_dockPopup.Hide
    End If
    
    'If m_systemTray.Dimensions.Left <> Me.ScaleWidth - m_systemTray.Image.Width Then
    'CalculateTaskbandDimensions
    'End If
    
    m_taskList.Update Me.hWnd
    m_menuBar.Update
    
    'm_taskList.PrintProcesses
    
    ResizeTaskBandIfNeeded m_taskList.Processes.Count
    
    SetPositionNewProcesses
    CopyAndSortVisibleList
    
    DrawTaskList
End Sub

Sub DrawTaskList()

    On Error GoTo Handler

    Dim thisProcess As Process

    Dim szDebugInfo As String

    m_TaskbarBackgroundGraphics.Clear
    m_TaskbandGraphics.Clear
    
    'm_TaskbarBackgroundGraphics.Exclude m_menuBar.Dimensions
    m_TaskbarBackgroundGraphics.Exclude RECTLtoF(m_menuBar.Dimensions)
    
    'm_TaskbarBackgroundGraphics.DrawImage m_backgroundImage, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    SliceHelper.DrawSlices m_slices, m_TaskbarBackgroundGraphics, Me
    
    m_TaskbarBackgroundGraphics.ResetExclusions
    
    If Not m_visibleList Is Nothing Then
        m_processIndex = 1

        For Each thisProcess In m_visibleList

            DrawProcess thisProcess
            m_processIndex = m_processIndex + 1
        Next

    End If
    
    m_TaskbarBackgroundGraphics.DrawImage m_menuBar.Image, CSng(m_menuBar.Dimensions.Left), CSng(m_menuBar.Dimensions.Top), CSng(m_menuBar.Dimensions.Width), CSng(m_menuBar.Dimensions.Height)

    
    'm_TaskbarBackgroundGraphics.DrawImage m_systemTray.Image, m_systemTray.Dimensions.Left, m_systemTray.Dimensions.Top, m_systemTray.Image.Width, m_systemTray.Image.Height
    m_TaskbarBackgroundGraphics.DrawImage m_TaskbandBitmap.Image, CSng(m_taskbandDimensions.Left), EFFECT_MARGIN, CSng(m_taskbandDimensions.Width), CSng(m_taskbandDimensions.Height)
    
    'm_TaskbarBackgroundGraphics.DrawString "File", FontHelper.AppDefaultFont, FontHelper.GetBlackBrush, CreatePointF(0, 0)
    m_clock.Draw
    
    MenuListHelper.DrawButton m_arrowSlices, m_arrowState, m_TaskbarBackgroundGraphics, m_arrowPlacement

    m_layeredWindowProperties.Update Me.hWnd, m_layeredWindowProperties.theDC

    Exit Sub

Handler:
    LogError Err.Number, "DrawTaskList(); " & Err.Description & vbCrLf & szDebugInfo, "winTaskBar"
End Sub

Private Sub ResetProcessPositions()

    Dim thisProcess      As Process

    For Each thisProcess In m_taskList.Processes
        thisProcess.X = -1
    Next

End Sub

Private Sub SetPositionNewProcesses()

    Dim thisProcess      As Process

    Dim thisProcessIndex As Long

    thisProcessIndex = m_taskList.Processes.Count

    For Each thisProcess In m_taskList.Processes

        If thisProcess.X = -1 Then
            thisProcess.X = (m_taskList.Processes.Count - thisProcessIndex) * (ACTUAL_ICON_SIZE + SEPERATOR_SIZE)
        End If
        
        thisProcessIndex = thisProcessIndex - 1
    Next

End Sub

Sub CopyAndSortVisibleList()

    Dim processTemp  As Process

    Dim lngX         As Long

    Dim lngY         As Long

    If m_taskList.Processes.Count > 0 Then
        Set m_visibleList = New Collection

        For Each processTemp In m_taskList.Processes
            m_visibleList.Add processTemp
        Next
    
        For lngX = 1 To m_visibleList.Count - 1
            For lngY = 1 To m_visibleList.Count - 1
            
                If m_visibleList(lngY).X > m_visibleList(lngY + 1).X Then
                    ' exchange the items
                    Set processTemp = m_visibleList(lngY)
                    m_visibleList.Remove lngY
                    
                    AddToCollectionAtPosition m_visibleList, processTemp, lngY + 1
                End If
                
            Next
        Next

    End If

End Sub

Private Function GetSelectedProcess(ByVal X As Single) As Process

    Dim thisItem As Process

    Dim startX   As Long

    Dim endX     As Long

    startX = m_taskbandDimensions.Left

    For Each thisItem In m_visibleList

        endX = startX + (ACTUAL_ICON_SIZE + SEPERATOR_SIZE)
        
        If X > startX And X < endX Then
            Set GetSelectedProcess = thisItem

            Exit For

        End If
        
        startX = endX
    Next

End Function

Private Function DrawProcess(ByRef thisProcess As Process)

    On Error GoTo Handler
    
    Dim indicatorPosition As gdiplus.RECTL

    indicatorPosition = CreateRectL(8, 8, thisProcess.X + (ACTUAL_ICON_SIZE) - 4, thisProcess.Y + (ACTUAL_ICON_SIZE - 4))
    
    If thisProcess Is Nothing Then

        Exit Function

    End If
    
    If thisProcess.Image Is Nothing Then

        Exit Function

    End If

    m_TaskbandGraphics.DrawImage thisProcess.Image, thisProcess.X, 0, ACTUAL_ICON_SIZE, ACTUAL_ICON_SIZE
    
    If thisProcess.Running Then
        DrawButton m_indicatorSlices, ButtonPressed, m_TaskbandGraphics, indicatorPosition
    ElseIf thisProcess.Flashing Then
        DrawButton m_indicatorSlices, ButtonNotice, m_TaskbandGraphics, indicatorPosition
    Else
        DrawButton m_indicatorSlices, ButtonUnpressed, m_TaskbandGraphics, indicatorPosition
    End If

    Exit Function

Handler:
    LogError Err.Number, "DrawProcess():" & Err.Description, "Taskbar"
End Function

Private Sub Timer1_Timer()
    Me.Width = Me.Width - 10 * Screen.TwipsPerPixelX
End Sub

Private Sub timUpdateClock_Timer()
    m_clock.Update
End Sub
