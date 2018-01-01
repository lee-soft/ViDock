VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Slice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : Slice
'    Project    : ViDock
'
'    Description: A managed segment of a bitmap. Describes a portion of a skin
'                 bitmap resource
'
'--------------------------------------------------------------------------------
Option Explicit

Public Anchor         As AnchorPointConstants

Public Image          As GDIPImage

Public X              As Long

Public Y              As Long

Public StretchX       As Boolean

Public StretchY       As Boolean

Public StretchMarginX As Long

Public StretchMarginY As Long

Public Identifer      As String

Public Property Get Height() As Long
    If Image Is Nothing Then
        Exit Property
    End If
    
    Height = Image.Height
End Property

Public Property Get Width() As Long
    If Image Is Nothing Then
        Exit Property
    End If

    Width = Image.Width
End Property

