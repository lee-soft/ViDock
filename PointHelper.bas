Attribute VB_Name = "PointHelper"
'--------------------------------------------------------------------------------
'    Component  : PointHelper
'    Project    : ViDock
'
'    Description: Provides functions that quickly create pointers and pointerfs
'
'--------------------------------------------------------------------------------
Option Explicit

Public Function CreatePoint(X As Long, Y As Long) As gdiplus.POINTL

    Dim newPoint As gdiplus.POINTL

    With newPoint
        .X = X
        .Y = Y
    End With

    CreatePoint = newPoint
End Function

Public Function CreatePointF(X As Long, Y As Long) As gdiplus.POINTF

    Dim newPoint As gdiplus.POINTF

    With newPoint
        .X = X
        .Y = Y
    End With

    CreatePointF = newPoint
End Function
