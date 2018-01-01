Attribute VB_Name = "RectHelper"
'--------------------------------------------------------------------------------
'    Component  : RectHelper
'    Project    : ViDock
'
'    Description: Contains helper functions for RECTs and RECTFs
'
'--------------------------------------------------------------------------------


Option Explicit

Public Function RECTWIDTH(ByRef srcRect As Win.RECT)
    RECTWIDTH = srcRect.Right - srcRect.Left
End Function

Public Function RECTHEIGHT(ByRef srcRect As Win.RECT)
    RECTHEIGHT = srcRect.Bottom - srcRect.Top
End Function

Public Function PrintRectF(ByRef srcRect As gdiplus.RECTF)
    Debug.Print "Top; " & srcRect.Top & vbCrLf & "Left; " & srcRect.Left & vbCrLf & "Height; " & srcRect.Height & vbCrLf & "Width; " & srcRect.Width
End Function

Public Function PrintRect(ByRef srcRect As RECT)
    Debug.Print "Top; " & srcRect.Top & vbCrLf & "Left; " & srcRect.Left & vbCrLf & "Bottom; " & srcRect.Bottom & vbCrLf & "Right; " & srcRect.Right
End Function

Public Function RECTtoF(ByRef srcRECTL As RECT) As gdiplus.RECTF
    RECTtoF = CreateRectF(CLng(srcRECTL.Left), CLng(srcRECTL.Top), CLng(srcRECTL.Bottom), CLng(srcRECTL.Right))
End Function

Public Function RECTFtoL(ByRef srcRect As gdiplus.RECTF) As gdiplus.RECTL
    RECTFtoL = CreateRectL(srcRect.Height, srcRect.Width, srcRect.Left, srcRect.Top)
End Function

Public Function RECTLtoF(ByRef srcRECTL As gdiplus.RECTL) As gdiplus.RECTF
    RECTLtoF = CreateRectF(CLng(srcRECTL.Left), CLng(srcRECTL.Top), CLng(srcRECTL.Height), CLng(srcRECTL.Width))
End Function

Public Function PointInsideOfRect(srcPoint As Win.POINTL, srcRect As Win.RECT) As Boolean

    PointInsideOfRect = False

    If srcPoint.Y > srcRect.Top And srcPoint.Y < srcRect.Bottom And srcPoint.X > srcRect.Left And srcPoint.X < srcRect.Right Then

        PointInsideOfRect = True
    End If

End Function

Public Function CreateRect(Left As Long, _
                           Top As Long, _
                           Bottom As Long, _
                           Right As Long) As RECT

    Dim newRect As RECT

    With newRect
        .Left = Left
        .Top = Top
        .Bottom = Bottom
        .Right = Right
    End With
    
    CreateRect = newRect
End Function

Public Function CreateRectL(ByVal lHeight As Long, _
                            ByVal lWidth As Long, _
                            ByVal lLeft As Long, _
                            ByVal lTop As Long) As gdiplus.RECTL

    With CreateRectL
        .Height = lHeight
        .Left = lLeft
        .Top = lTop
        .Width = lWidth
    End With

End Function

Public Function CreateRectF(Left As Single, _
                            Top As Single, _
                            Height As Single, _
                            Width As Single) As gdiplus.RECTF

    Dim newRectF As gdiplus.RECTF

    With newRectF
        .Left = Left
        .Top = Top
        .Height = Height
        .Width = Width
    End With
    
    CreateRectF = newRectF
End Function

