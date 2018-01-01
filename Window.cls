VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : Window
'    Project    : ViDock
'
'    Description: The Window object is an object orientated wrapper of
'                 a Window in Windows the operating system. myWindow.caption
'                 would be "Notepad" for example.
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

'local variable(s) to hold property value(s)
Private mvarhWnd     As Long 'local copy

Private mvarCaption  As String 'local copy

Private mvarFlashing As Boolean 'local copy

Private mvarParent   As Process 'local copy

Private mvarBitmap   As GDIPBitmap

Public Function IsHung() As Boolean

    Dim lResult As Long

    Dim lReturn As Long
    
    lReturn = SendMessageTimeout( _
        mvarhWnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG Or SMTO_BLOCK, 1000, lResult)
                     
    If lReturn Then
        IsHung = False

        Exit Function

    End If
    
    IsHung = True
End Function

Public Property Get Image() As GDIPImage
    Set Image = mvarBitmap.Image
End Property

Public Function IsMaximized() As Boolean
    IsMaximized = IIf(IsZoomed(mvarhWnd) = APITRUE, True, False)
End Function

Public Function isMinimized() As Boolean
    isMinimized = IIf(IsIconic(mvarhWnd) = APITRUE, True, False)
End Function

Public Function UpdateImage()

    On Error GoTo Handler
    
    If Me.IsHung Then

        Exit Function

    End If

    Dim g             As New GDIPGraphics

    Dim bm            As New GDIPBitmap

    Dim bmTemp        As New GDIPBitmap

    Dim gTemp         As New GDIPGraphics

    Dim hdc           As Long

    Dim rc            As Win.RECT

    Dim hBitmap       As Long

    Dim newWidth      As Long

    Dim newHeight     As Long

    Dim aspect_width  As Single

    Dim aspect_height As Single

    mvarBitmap.Dispose
    
    GetWindowRect mvarhWnd, rc
    
    aspect_width = RECTWIDTH(rc) / RECTHEIGHT(rc)
    aspect_height = RECTHEIGHT(rc) / RECTWIDTH(rc)
    
    newWidth = WINDOW_IMAGE_HEIGHT * aspect_width
    newHeight = WINDOW_IMAGE_WIDTH * aspect_height
    
    If newWidth > WINDOW_IMAGE_WIDTH Or WINDOW_IMAGE_WIDTH < 0 Then
        newWidth = WINDOW_IMAGE_WIDTH
    End If
    
    If newHeight > WINDOW_IMAGE_HEIGHT Then
        newHeight = WINDOW_IMAGE_HEIGHT
    End If
    
    bm.CreateFromSizeFormat RECTWIDTH(rc), RECTHEIGHT(rc), GDIPlusWrapper.Format32bppArgb
    
    g.FromImage bm.Image

    hdc = g.GetHDC
    
    RepaintWindow mvarhWnd

    If PrintWindow(mvarhWnd, hdc, 0) = 0 Then
        Debug.Print "UpdateImage failed"
    End If
    
    g.ReleaseHDC hdc
    g.flush FlushIntentionFlush
    
    hBitmap = bm.hBitmap(vbBlack)
    bmTemp.CreateFromHBITMAP hBitmap, 0
    DeleteObject hBitmap
    
    mvarBitmap.CreateFromSizeFormat newWidth, newHeight, GDIPlusWrapper.Format32bppArgb
    gTemp.FromImage mvarBitmap.Image
    gTemp.SmoothingMode = SmoothingModeHighQuality
    gTemp.InterpolationMode = InterpolationModeHighQualityBicubic
    gTemp.PixelOffsetMode = PixelOffsetModeHighQuality
    
    gTemp.DrawImage bmTemp.Image, 0, 0, CSng(newWidth), CSng(newHeight)

Handler:
End Function

Public Property Get WindowState() As FormWindowStateConstants
    
    If isMinimized Then
        WindowState = vbMinimized
    ElseIf IsMaximized Then
        WindowState = vbMaximized
    Else
        WindowState = vbNormal
    End If
    
End Property

Public Property Set Parent(ByVal vData As Process)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.Parent = Form1
    Set mvarParent = vData
End Property

Public Property Get Parent() As Process
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Parent
    Set Parent = mvarParent
End Property

Public Property Let Flashing(ByVal vData As Boolean)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Flashing = 5
    mvarFlashing = vData
End Property

Public Property Get Flashing() As Boolean
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Flashing
    Flashing = mvarFlashing
End Property

Public Property Let Caption(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Caption = 5
    mvarCaption = vData
End Property

Public Property Get Caption() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property

Public Property Let hWnd(ByVal vData As Long)

    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.hWnd = 5
    If vData = 0 Then
        LogError 299, "hWnd->Invalid hWnd", "Class::Window"

        Exit Property

    End If

    mvarhWnd = vData
    UpdateWindowText
End Property

Public Function UpdateWindowText()

    If mvarhWnd = 0 Then Exit Function

    If Me.IsHung Then Exit Function

    Dim lReturn As Long
    
    mvarCaption = Space$(256)
    lReturn = GetWindowText(hWnd, mvarCaption, Len(mvarCaption))
    
    If lReturn Then
        mvarCaption = Left$(mvarCaption, lReturn)
    End If

End Function

Public Property Get hWnd() As Long
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.hWnd
    hWnd = mvarhWnd
End Property

Private Sub Class_Initialize()
    Set mvarBitmap = New GDIPBitmap
End Sub

