VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DockSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PinnedApplications As Collection

Public LoadDefaults       As Boolean

Public ClassBlackList     As SimpleFileList

Private m_settingsDoc     As DOMDocument

Public Function Comit()

    Dim settingsXML As IXMLDOMElement

    Dim generalXML  As IXMLDOMElement

    'Dim navigationXML As IXMLDOMElement

    Set m_settingsDoc = New DOMDocument
    
    Set settingsXML = m_settingsDoc.createElement("settings")
    Set generalXML = m_settingsDoc.createElement("general")
    
    m_settingsDoc.appendChild settingsXML
    settingsXML.appendChild generalXML

    If Not DumpPinnedProcesses(m_settingsDoc, settingsXML, PinnedApplications) Then

        Exit Function

    End If
    
    m_settingsDoc.save App.Path & "\settings.xml"
End Function

Private Sub Class_Initialize()
    Set m_settingsDoc = New DOMDocument
    Set PinnedApplications = New Collection
    Set ClassBlackList = New SimpleFileList
    
    Dim settingsXML As IXMLDOMElement

    ClassBlackList.Initialize App.Path & "\classblacklist.lst"

    If Not m_settingsDoc.Load(App.Path & "\settings.xml") Then
        LoadDefaults = True

        Exit Sub

    End If
    
    Set settingsXML = m_settingsDoc.firstChild
    
    '    If Not IsNull(settingsXML.selectSingleNode("//pinned_programs")) Then
    '        Set programsXML = settingsXML.selectSingleNode("//pinned_programs")
    '        If Not IsNull(programsXML) Then ReadPinnedPrograms programsXML, Me.PinnedApplications
    '    End If
End Sub
