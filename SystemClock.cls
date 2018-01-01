VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : SystemClock
'    Project    : ViDock
'
'    Description: [type_description_here]
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_dateTime As String

Private m_Point    As gdiplus.POINTF

Private m_width    As Long

Private m_graphics As GDIPGraphics

Public Event onPopup(ByVal szText As String)

Public Event onMouseLeaves()

Public Sub Draw()
    m_graphics.DrawString m_dateTime, AppDefaultFont, GetBlackBrush, m_Point
End Sub

Public Sub MouseLeft()
    RaiseEvent onMouseLeaves
End Sub

Public Sub MouseDown(Button As Integer, X As Single, Y As Single)

End Sub

Public Sub MouseMove(Button As Integer, X As Single, Y As Single)
    RaiseEvent onPopup(Format(Now, "dddd, mmmm, dd - hh:mm:ss AM/PM"))
End Sub

Public Sub Update()
    m_dateTime = Format(Now(), "dddd hh:mm AM/PM")
End Sub

Public Sub Initialize(ByRef theGraphics As GDIPGraphics)
    Set m_graphics = theGraphics
    Update
    
    m_width = m_graphics.MeasureString(m_dateTime, AppDefaultFont).Width
End Sub

Public Property Get Width() As Long
    Width = m_width
End Property

Public Property Let X(newX As Single)
    m_Point.X = newX
End Property

Public Property Get X() As Single
    X = m_Point.X
End Property

Public Property Let Y(newY As Single)
    m_Point.Y = newY
End Property

