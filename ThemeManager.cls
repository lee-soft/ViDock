VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThemeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_themeDoc As DOMDocument

Private Sub Class_Initialize()


    Set m_themeDoc = New DOMDocument
    If m_themeDoc.Load(App.Path & "\resources\theme.xml") = False Then Exit Sub

    ProcessXMLElements m_themeDoc.firstChild
End Sub

Private Sub ProcessXMLElements(ByRef xmlRoot As IXMLDOMElement)
    
Dim thisIncludedDoc As DOMDocument
Dim thisChild As IXMLDOMElement
    
    For Each thisChild In xmlRoot.childNodes
        If thisChild.tagName = "xi:include" Then
            
            Set thisIncludedDoc = New DOMDocument
            
            If thisIncludedDoc.Load(App.Path & "\resources\" & thisChild.getAttribute("href")) Then
                ProcessXMLElements thisIncludedDoc.firstChild
            End If
        Else
            Debug.Print thisChild.tagName
        End If
    Next
    
End Sub
