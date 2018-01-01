VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrayButtonInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : TrayButtonInfo
'    Project    : ViDock
'
'    Description: Holds all the data for each button in the tray icon area
'
'--------------------------------------------------------------------------------


Option Explicit

Public Tooltip          As String

Public hIcon            As Long

Public hWnd             As Long

Public uCallbackMessage As Long

Public uID              As Long

'Public IconPosition As RECT
Public Visible          As Boolean

Public Property Get App_Identifier() As String
    App_Identifier = "but_" & hWnd
End Property
