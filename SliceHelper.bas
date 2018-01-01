Attribute VB_Name = "SliceHelper"
'--------------------------------------------------------------------------------
'    Component  : SliceHelper
'    Project    : ViDock
'
'    Description: Contains helper functions for managing slices
'
'--------------------------------------------------------------------------------
Option Explicit

Public Enum AnchorPointConstants

    apTopLeft = 1
    apBottomLeft = 2
    apBottomRight = 3
    apTopRight = 4

End Enum

Public Function FindBorderWidth(theSlices As Collection) As Long
    
    Dim thisSlice As Slice

    For Each thisSlice In theSlices

        If thisSlice.Anchor = apTopLeft And thisSlice.StretchX = False And thisSlice.StretchY = True Then
            FindBorderWidth = thisSlice.Width

            Exit For

        End If

    Next
    
End Function

Public Function FindBorderHeight(theSlices As Collection) As Long
    
    Dim thisSlice As Slice

    For Each thisSlice In theSlices

        If thisSlice.Anchor = apTopLeft And thisSlice.StretchX = True And thisSlice.StretchY = False Then
            FindBorderHeight = thisSlice.Height

            Exit For

        End If

    Next
    
End Function

Public Function AnchorPointTextToLong(ByVal szText As String) As AnchorPointConstants

    Select Case LCase(szText)
    
        Case "lt", "tl"
            AnchorPointTextToLong = apTopLeft
        
        Case "rt", "tr"
            AnchorPointTextToLong = apTopRight
        
        Case "bl", "lb"
            AnchorPointTextToLong = apBottomLeft
        
        Case "rb", "br"
            AnchorPointTextToLong = apBottomRight
        
    End Select

End Function

Public Function DrawSlicesToTargetF(ByRef slices As Collection, _
                                    ByRef Graphics As GDIPGraphics, _
                                    theTarget As gdiplus.RECTF)

    Debug.Print "DrawSlicesToTargetF"

    Dim thisSlice As Slice

    For Each thisSlice In slices

        Select Case thisSlice.Anchor
        
            Case AnchorPointConstants.apTopLeft

                If thisSlice.StretchX And Not thisSlice.StretchY Then
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, theTarget.Width - thisSlice.StretchMarginX, thisSlice.Height
                                                    
                ElseIf thisSlice.StretchY And Not thisSlice.StretchX Then
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, theTarget.Height - thisSlice.StretchMarginY
                                                    
                ElseIf thisSlice.StretchX And thisSlice.StretchY Then
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, theTarget.Width - thisSlice.StretchMarginX, theTarget.Height - thisSlice.StretchMarginY
                Else
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If
        
            Case AnchorPointConstants.apTopRight

                If thisSlice.StretchY Then
                    Graphics.DrawImage thisSlice.Image, (theTarget.Left + theTarget.Width) - thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, theTarget.Height - thisSlice.StretchMarginY
                Else
                    Graphics.DrawImage thisSlice.Image, (theTarget.Left + theTarget.Width) - thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If
        
            Case AnchorPointConstants.apBottomLeft

                If thisSlice.StretchX Then
                    'Graphics.DrawImage thisSlice.Image, thisSlice.X, theForm.ScaleHeight - thisSlice.y, theForm.ScaleWidth - thisSlice.StretchMarginX, thisSlice.Height
            
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, (theTarget.Top + theTarget.Height) - thisSlice.Y, theTarget.Width - thisSlice.StretchMarginX, thisSlice.Height
                Else
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, (theTarget.Top + theTarget.Height) - thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If
        
            Case AnchorPointConstants.apBottomRight

                If thisSlice.StretchX Then
            
                Else
                    Graphics.DrawImage thisSlice.Image, (theTarget.Left + theTarget.Width) - thisSlice.X, (theTarget.Top + theTarget.Height) - thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If

        End Select

    Next

End Function

Public Function DrawSlicesToTarget(ByRef slices As Collection, _
                                   ByRef Graphics As GDIPGraphics, _
                                   theTarget As gdiplus.RECTL)

    Dim thisSlice As Slice

    For Each thisSlice In slices

        Select Case thisSlice.Anchor
        
            Case AnchorPointConstants.apTopLeft

                If thisSlice.StretchX And Not thisSlice.StretchY Then
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, theTarget.Width - thisSlice.StretchMarginX, thisSlice.Height
                                                    
                ElseIf thisSlice.StretchY And Not thisSlice.StretchX Then
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, theTarget.Height - thisSlice.StretchMarginY
                                                    
                ElseIf thisSlice.StretchX And thisSlice.StretchY Then
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, theTarget.Width - thisSlice.StretchMarginX, theTarget.Height - thisSlice.StretchMarginY
                Else
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If
        
            Case AnchorPointConstants.apTopRight

                If thisSlice.StretchY Then
                    Graphics.DrawImage thisSlice.Image, (theTarget.Left + theTarget.Width) - thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, theTarget.Height - thisSlice.StretchMarginY
                Else
                    Graphics.DrawImage thisSlice.Image, (theTarget.Left + theTarget.Width) - thisSlice.X, theTarget.Top + thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If
        
            Case AnchorPointConstants.apBottomLeft

                If thisSlice.StretchX Then
                    'Graphics.DrawImage thisSlice.Image, thisSlice.X, theForm.ScaleHeight - thisSlice.y, theForm.ScaleWidth - thisSlice.StretchMarginX, thisSlice.Height
            
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, (theTarget.Top + theTarget.Height) - thisSlice.Y, theTarget.Width - thisSlice.StretchMarginX, thisSlice.Height
                Else
                    Graphics.DrawImage thisSlice.Image, theTarget.Left + thisSlice.X, (theTarget.Top + theTarget.Height) - thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If
        
            Case AnchorPointConstants.apBottomRight

                If thisSlice.StretchX Then
            
                Else
                    Graphics.DrawImage thisSlice.Image, (theTarget.Left + theTarget.Width) - thisSlice.X, (theTarget.Top + theTarget.Height) - thisSlice.Y, thisSlice.Width, thisSlice.Height
                End If

        End Select

    Next

End Function

Public Function DrawSlices(ByRef slices As Collection, _
                           ByRef Graphics As GDIPGraphics, _
                           ByRef theForm As Form)

    Dim thisSlice As Slice

    For Each thisSlice In slices

        If thisSlice.Identifer = vbNullString Then

            Select Case thisSlice.Anchor
            
                Case AnchorPointConstants.apTopLeft

                    If thisSlice.StretchX And Not thisSlice.StretchY Then
                        Graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, theForm.ScaleWidth - thisSlice.StretchMarginX, thisSlice.Height
                    ElseIf thisSlice.StretchY And Not thisSlice.StretchX Then
                        Graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, thisSlice.Width, theForm.ScaleHeight - thisSlice.StretchMarginY
                    ElseIf thisSlice.StretchX And thisSlice.StretchY Then
                        Graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, theForm.ScaleWidth - thisSlice.StretchMarginX, theForm.ScaleHeight - thisSlice.StretchMarginY
                    Else
                        Graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, thisSlice.Width, thisSlice.Height
                    End If
            
                Case AnchorPointConstants.apTopRight

                    If thisSlice.StretchY Then
                        Graphics.DrawImage thisSlice.Image, theForm.ScaleWidth - thisSlice.X, thisSlice.Y, thisSlice.Width, theForm.ScaleHeight - thisSlice.StretchMarginY
                    Else
                        Graphics.DrawImage thisSlice.Image, theForm.ScaleWidth - thisSlice.X, thisSlice.Y, thisSlice.Width, thisSlice.Height
                    End If
            
                Case AnchorPointConstants.apBottomLeft

                    If thisSlice.StretchX Then
                        Graphics.DrawImage thisSlice.Image, thisSlice.X, theForm.ScaleHeight - thisSlice.Y, theForm.ScaleWidth - thisSlice.StretchMarginX, thisSlice.Height
                    Else
                        Graphics.DrawImage thisSlice.Image, thisSlice.X, theForm.ScaleHeight - thisSlice.Y, thisSlice.Width, thisSlice.Height
                    End If
            
                Case AnchorPointConstants.apBottomRight

                    If thisSlice.StretchX Then
                
                    Else
                        Graphics.DrawImage thisSlice.Image, theForm.ScaleWidth - thisSlice.X, theForm.ScaleHeight - thisSlice.Y, thisSlice.Width, thisSlice.Height
                    End If
                
                    'Graphics.DrawImage thisSlice.Image, theForm.ScaleWidth - thisSlice.X, theForm.ScaleHeight - thisSlice.Y, thisSlice.Width, thisSlice.Height
    
            End Select

        End If

    Next

End Function

Public Function CreateSlicesFromXML(ByVal szXmlElementName As String, _
                                    Optional ByRef slicesImage As GDIPImage) As Collection
    
    Dim slicesXMLElement As IXMLDOMElement

    Set slicesXMLElement = ThemeHelper.GetSliceDefinition(szXmlElementName)
    
    If slicesXMLElement Is Nothing Then
        LogError 0, "CreateSlicesFromXML", "SliceHelper", "Warning no xml definition found for: " & szXmlElementName

        Exit Function

    End If
    
    Set slicesImage = New GDIPImage
    slicesImage.FromFile App.Path & "\resources\" & slicesXMLElement.getAttribute("src")
    
    Set CreateSlicesFromXML = CreateSlicesFromXMLElement(slicesXMLElement, slicesImage)
End Function

Public Function CreateSlicesFromXMLElement(ByRef slicesXML As IXMLDOMElement, _
                                           ByRef slicesImage As GDIPImage) As Collection

    Dim thisXMLSlice      As IXMLDOMElement

    Dim thisSlice         As Slice

    Dim sliceWidth        As Long

    Dim sliceHeight       As Long

    Dim slices            As Collection

    Dim topSliceWidth     As Long

    Dim stretchDimensions As String

    Dim StretchMarginX    As Long

    Dim StretchMarginY    As Long

    Dim yOverflow         As Long

    Dim xOverflow         As Long

    topSliceWidth = -1

    Set slices = New Collection
    Set CreateSlicesFromXMLElement = slices
    
    For Each thisXMLSlice In slicesXML.childNodes

        If thisXMLSlice.tagName = "slice" Then
            StretchMarginX = -1
            StretchMarginY = -1

            Set thisSlice = New Slice
            
            If Not IsNull(thisXMLSlice.getAttribute("anchor")) Then
                thisSlice.Anchor = AnchorPointTextToLong(thisXMLSlice.getAttribute("anchor"))
            End If

            If Not IsNull(thisXMLSlice.getAttribute("stretch")) Then
                stretchDimensions = CStr(thisXMLSlice.getAttribute("stretch"))
                
                If InStr(stretchDimensions, "x") > 0 Then
                    thisSlice.StretchX = True
                End If
                
                If InStr(stretchDimensions, "y") > 0 Then
                    thisSlice.StretchY = True
                End If
            End If
            
            If Not IsNull(thisXMLSlice.getAttribute("x-margin")) Then
                StretchMarginX = CLng(thisXMLSlice.getAttribute("x-margin"))
            End If
            
            If Not IsNull(thisXMLSlice.getAttribute("y-margin")) Then
                StretchMarginY = CLng(thisXMLSlice.getAttribute("y-margin"))
            End If

            If Not IsNull(thisXMLSlice.getAttribute("id")) Then
                thisSlice.Identifer = CStr(thisXMLSlice.getAttribute("id"))
            End If

            thisSlice.X = CLng(thisXMLSlice.getAttribute("x"))
            thisSlice.Y = CLng(thisXMLSlice.getAttribute("y"))
            
            sliceWidth = CLng(thisXMLSlice.getAttribute("width"))
            sliceHeight = CLng(thisXMLSlice.getAttribute("height"))
            
            If Not IsNull(thisXMLSlice.getAttribute("x-overflow")) Then
                xOverflow = CLng(thisXMLSlice.getAttribute("x-overflow"))
            End If
            
            If Not IsNull(thisXMLSlice.getAttribute("y-overflow")) Then
                yOverflow = CLng(thisXMLSlice.getAttribute("y-overflow"))
            End If

            Set thisSlice.Image = CreateNewImageFromSection(slicesImage, CreateRectL(sliceHeight, sliceWidth, thisSlice.X, thisSlice.Y))
            If thisSlice.Image Is Nothing Then
                MsgBox "Error!"
            End If
            
            
            thisSlice.Y = thisSlice.Y + yOverflow
            thisSlice.X = thisSlice.X + xOverflow
            
            If thisSlice.Anchor = apTopRight Then
                thisSlice.X = slicesImage.Width - thisSlice.X
            ElseIf thisSlice.Anchor = apBottomLeft Then
                thisSlice.Y = slicesImage.Height - thisSlice.Y
            ElseIf thisSlice.Anchor = apBottomRight Then
                thisSlice.X = slicesImage.Width - thisSlice.X
                thisSlice.Y = slicesImage.Height - thisSlice.Y
            End If
            
            If StretchMarginX = -1 Then
                thisSlice.StretchMarginX = slicesImage.Width - sliceWidth
            Else
                thisSlice.StretchMarginX = StretchMarginX
            End If
            
            If StretchMarginY = -1 Then
                If slicesImage.Height = 74 Then

                End If
            
                thisSlice.StretchMarginY = slicesImage.Height - sliceHeight
            Else
                thisSlice.StretchMarginY = StretchMarginY
            End If
                                                                
            If thisSlice.Identifer <> vbNullString Then
                slices.Add thisSlice, thisSlice.Identifer
            Else
                slices.Add thisSlice
            End If
        End If
        
    Next
    
End Function

Public Function CreateNewImageFromSection(ByRef sourceImage As GDIPImage, _
                                          sourceSection As gdiplus.RECTL) As GDIPImage

    Dim returnBitmap As GDIPBitmap

    Dim tempGraphics As GDIPGraphics

    Set returnBitmap = New GDIPBitmap
    
    'returnBitmap.CreateFromSize sourceSection.Width, sourceSection.Height
    returnBitmap.CreateFromSizeFormat sourceSection.Width, sourceSection.Height, GDIPlusWrapper.Format32bppArgb
    
    Set tempGraphics = New GDIPGraphics
    tempGraphics.FromImage returnBitmap.Image
    tempGraphics.Clear
    
    tempGraphics.DrawImageRect sourceImage, 0, 0, sourceSection.Width, sourceSection.Height, sourceSection.Left, sourceSection.Top

    Set tempGraphics = Nothing
    Set CreateNewImageFromSection = returnBitmap.Image.Clone
End Function
