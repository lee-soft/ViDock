Attribute VB_Name = "ThemeHelper"
'--------------------------------------------------------------------------------
'    Component  : ThemeHelper
'    Project    : ViDock
'
'    Description: Parses the theme XML files
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_themeDoc                 As DOMDocument

Public ListWindowClipped           As Collection

Public ListWindowClippedImage      As GDIPImage

Public ListWindow                  As Collection

Public ListWindowImage             As GDIPImage

Public ListWindowTextMargin        As MARGIN

Public ListWindowTextClippedMargin As MARGIN

Public Buttons                     As Collection

Public SliceDefinitions            As Collection

Public Margins                     As Collection

Public Colour1                     As Colour

Public Colour2                     As Colour

Public Seperator                   As GDIPImage

Public FontSize                    As Single

Public Font                        As GDIPFont

Public FontBold                    As GDIPFont

Function Initialize() As Boolean
    'On Error GoTo Handler

    Set Buttons = New Collection
    Set SliceDefinitions = New Collection
    Set Margins = New Collection
    
    Set m_themeDoc = New DOMDocument

    If m_themeDoc.Load(App.Path & "\resources\theme.xml") = False Then Exit Function

    Set Seperator = New GDIPImage: Seperator.FromFile App.Path & "\resources\separator.png"

    ProcessXMLElements m_themeDoc.firstChild
    GenerateFonts
    
    Set ListWindowClippedImage = New GDIPImage
    Set ListWindowImage = New GDIPImage
    
    Set ListWindowClipped = SliceHelper.CreateSlicesFromXML("list_window_clipped", ListWindowClippedImage)
    Set ListWindow = SliceHelper.CreateSlicesFromXML("list_window", ListWindowImage)
    
    Set ListWindowTextMargin = ThemeHelper.GetMargin("list_window_text")
    Set ListWindowTextClippedMargin = ThemeHelper.GetMargin("list_window_text_clipped")
    
    Initialize = True

    Exit Function

Handler:
    LogError 0, "Initialize", "ThemeHelper", Err.Description
End Function

Public Function GetMargin(ByVal marginId As String) As MARGIN

    If Not ExistInCol(Margins, marginId) Then Exit Function
    Set GetMargin = Margins(marginId)
End Function

Public Function GetButton(ByVal buttonId As String) As IXMLDOMElement

    If Not ExistInCol(Buttons, buttonId) Then Exit Function
    Set GetButton = Buttons(buttonId)
End Function

Public Function GetSliceDefinition(ByVal sliceId As String) As IXMLDOMElement

    If Not ExistInCol(SliceDefinitions, sliceId) Then Exit Function
    Set GetSliceDefinition = SliceDefinitions(sliceId)
End Function

Private Sub ProcessFontXMLElement(ByRef xmlRoot As IXMLDOMElement)

    Dim fontFace As String

    If Not xmlRoot.tagName = "font" Then Exit Sub
    
    If Not IsNull(xmlRoot.getAttribute("size")) Then
        FontSize = xmlRoot.getAttribute("size")
    Else
        FontSize = 15
    End If
    
    If Font Is Nothing Then
    
        If Not IsNull(xmlRoot.getAttribute("face")) Then
            fontFace = xmlRoot.getAttribute("face")
        End If
        
        If Not IsNull(xmlRoot.getAttribute("colour1")) Then
            Set Colour1 = New Colour
            Colour1.SetColourByHex xmlRoot.getAttribute("colour1")
        End If
    
        If Not IsNull(xmlRoot.getAttribute("colour2")) Then
            Set Colour2 = New Colour
            Colour2.SetColourByHex xmlRoot.getAttribute("colour2")
        End If
        
        Dim thisFamily As GDIPFontFamily
        Set thisFamily = CreateFontFamily(fontFace)
        
        Set Font = New GDIPFont
        Font.Constructor thisFamily, FontSize, FontStyleRegular
        
        Set FontBold = New GDIPFont
        FontBold.Constructor thisFamily, FontSize, FontStyleBold
        
    End If

End Sub

Private Sub ProcessMarginXMLElements(ByRef xmlRoot As IXMLDOMElement)

    Dim thisMargin   As MARGIN

    Dim thisMarginId As String

    Dim thisChild    As IXMLDOMElement

    For Each thisChild In xmlRoot.childNodes

        If LCase(thisChild.tagName) = "margin" Then
        
            If Not IsNull(thisChild.getAttribute("id")) Then
                thisMarginId = thisChild.getAttribute("id")
                
            End If
        
            If Not thisMarginId = vbNullString Then
                Set thisMargin = New MARGIN
                Margins.Add thisMargin, thisMarginId
                
                With thisMargin
                
                    If Not IsNull(thisChild.getAttribute("height")) Then .Height = thisChild.getAttribute("height")
                        
                    If Not IsNull(thisChild.getAttribute("width")) Then .Width = thisChild.getAttribute("width")
                        
                    If Not IsNull(thisChild.getAttribute("x-overflow")) Then .X_Overflow = thisChild.getAttribute("x-overflow")
                        
                    If Not IsNull(thisChild.getAttribute("y-overflow")) Then .Y_Overflow = thisChild.getAttribute("y-overflow")
                        
                    If Not IsNull(thisChild.getAttribute("x")) Then .X = thisChild.getAttribute("x")
                        
                    If Not IsNull(thisChild.getAttribute("y")) Then .Y = thisChild.getAttribute("y")

                End With
                
            End If

        End If

    Next

End Sub

Private Sub GenerateFonts()

    Dim fontFamilies()        As GDIPFontFamily

    Dim numberReturned        As Long

    Dim privateFontCollection As GDIPPrivateFC

    Dim privateFontFamily     As GDIPFontFamily

    'Get Regular Font
    Set privateFontCollection = New GDIPPrivateFC
    Set privateFontFamily = New GDIPFontFamily

    privateFontCollection.AddFontFile (App.Path & "\resources\font.ttf")
    privateFontCollection.FontCollection.GetFamilies 1, fontFamilies, numberReturned

    If numberReturned > 0 Then
    
        Set privateFontFamily = fontFamilies(0)
         
        Set Font = New GDIPFont
        Font.Constructor privateFontFamily, FontSize, FontStyleRegular
        
        Set FontBold = New GDIPFont
        FontBold.Constructor privateFontFamily, FontSize, FontStyleBold
    End If

    'Get Bold Font
    Set privateFontCollection = New GDIPPrivateFC
    Set privateFontFamily = New GDIPFontFamily

    privateFontCollection.AddFontFile (App.Path & "\resources\font-bold.ttf")
    privateFontCollection.FontCollection.GetFamilies 1, fontFamilies, numberReturned
    
    If numberReturned > 0 Then
    
        Set privateFontFamily = fontFamilies(0)
         
        Set FontBold = New GDIPFont
        FontBold.Constructor privateFontFamily, FontSize, FontStyleRegular
    End If

    'SetStatusHelper m_privateFontCollection.AddFontFile(App.Path & "\resources\font-bold.ttf")

    'If Not IsNull(xmlRoot.getAttribute("face")) Then
    '    DefaultFace = xmlRoot.getAttribute("face")
    'End If
    
    'If Not IsNull(xmlRoot.getAttribute("colour1")) Then
    '    Set Colour1 = New Colour
    '    Colour1.SetColourByHex xmlRoot.getAttribute("colour1")

    'End If

    'If Not IsNull(xmlRoot.getAttribute("colour2")) Then
    '    Set Colour2 = New Colour
    '    Colour2.SetColourByHex xmlRoot.getAttribute("colour2")
    'End If
    
    'If Not IsNull(xmlRoot.getAttribute("size")) Then
    '    FontSize = xmlRoot.getAttribute("size")
    'Else
    '    FontSize = 17
    'End If

End Sub

Private Sub ProcessXMLElements(ByRef xmlRoot As IXMLDOMElement)
    
    Dim thisIncludedDoc As DOMDocument

    Dim thisChild       As IXMLDOMElement

    Dim thisElementID   As String

    If Not IsNull(xmlRoot.getAttribute("id")) Then
        thisElementID = xmlRoot.getAttribute("id")
    End If
    
    Debug.Print "Processing:: " & thisElementID
    
    Select Case LCase(xmlRoot.tagName)
    
        Case "button"

            If Not thisElementID = vbNullString Then Buttons.Add xmlRoot, thisElementID
    
        Case "slice_index"
            Debug.Print "Adding slice_index: " & thisElementID
        
            If Not thisElementID = vbNullString Then SliceDefinitions.Add xmlRoot, thisElementID
        
        Case "margins"
            ProcessMarginXMLElements xmlRoot
    
    End Select
    
    If LCase(xmlRoot.tagName) = "theme" Then
    
        For Each thisChild In xmlRoot.childNodes

            Select Case thisChild.tagName
            
                Case "xi:include"
                    Set thisIncludedDoc = New DOMDocument
                
                    If thisIncludedDoc.Load(App.Path & "\resources\" & thisChild.getAttribute("href")) Then
                        'For Each selectedChild In thisIncludedDoc.childNodes
                        ProcessXMLElements thisIncludedDoc.firstChild
                        'Next
                    
                    Else
                        LogError 0, "ProcessXMLElements", "ThemeHelper", "Failure processing:: " & thisChild.getAttribute("href")
                    End If
                
                Case "font"
                    ProcessFontXMLElement thisChild
                
            End Select

        Next
    
    End If
    
End Sub

