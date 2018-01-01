VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Menu As Long

Public Property Get Handle()
    Handle = m_Menu
End Property

Public Function ShowMenu(Optional hWnd As Long, _
                         Optional X As Long = 0, _
                         Optional y As Long = 0) As Long

    Dim CurrentCursorPos As Win.POINTL
    
    If X = 0 And y = 0 Then
        GetCursorPos CurrentCursorPos
        
        X = CurrentCursorPos.X
        y = CurrentCursorPos.y
    End If

    ShowMenu = TrackPopupMenu(m_Menu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_LEFTBUTTON, X, y, 0, hWnd, ByVal 0&)
End Function

Public Function GetMenuText(Command As Long)

    Dim itemLength   As Long

    Dim bufferString As String

    bufferString = String(1024, Chr(0))

    itemLength = GetMenuString(m_Menu, Command, bufferString, Len(bufferString), MF_BYCOMMAND)
    bufferString = Mid(bufferString, 1, itemLength)

    GetMenuText = bufferString

End Function

Public Function EditItem(Id As Long, strNewCaption As String)
    ModifyMenu m_Menu, Id, MF_STRING, ByVal Id, strNewCaption
End Function

Public Function AddItem(Id As Long, strItem As String)
    AppendMenu m_Menu, MF_STRING, ByVal Id, strItem
End Function

Public Function AddSeperater()
    AppendMenu m_Menu, MF_SEPARATOR, 0, 0
End Function

Private Sub Class_Initialize()
    'create menu
    m_Menu = CreatePopupMenu
End Sub

Private Sub Class_Terminate()
    DestroyMenu m_Menu
End Sub
