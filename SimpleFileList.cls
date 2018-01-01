VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : SimpleFileList
'    Project    : ViDock
'
'    Description: A initalizes a list from a file with each item seperated
'                 by a new line
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_list As Collection

Public Function Exists(szItem As String) As Boolean
    Exists = ExistInCol(m_list, szItem)
End Function

Public Function Initialize(szPath As String) As Boolean

    On Error GoTo Handler

    Dim ff       As Long

    Dim thisLine As String
    
    ff = FreeFile()
    
    Set m_list = New Collection

    Open szPath For Input As #ff

    Do While Not EOF(ff)
        Line Input #ff, thisLine
            
        If Not ExistInCol(m_list, thisLine) Then
            m_list.Add thisLine, thisLine
        End If

    Loop

    Close #ff

    Initialize = True
Handler:
End Function

