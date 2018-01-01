VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : MenuItem
'    Project    : ViDock
'
'    Description: Represents an OS file menu item.
'                 IE: "File" on Notepad.
'
'--------------------------------------------------------------------------------
Option Explicit

Public Width    As Single

Public Left     As Single

Public Caption  As String

Public Children As Collection

Public itemID   As Long

Public hWnd     As Long

Private Sub Class_Initialize()
    Set Children = New Collection
End Sub

