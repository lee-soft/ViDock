Attribute VB_Name = "GDIHelper"
Option Explicit

Public Type ARGB
    R As Byte
    g As Byte
    b As Byte
    a As Byte
End Type

Public Sub Long2ARGB(ByVal LongARGB As Long, ByRef ARGB As ARGB)
    WinAPIHelper.CopyMemory ARGB, LongARGB, 4
End Sub

Public Sub ARGB2Long(ByRef ARGB As ARGB, ByRef theLong As Long)
    WinAPIHelper.CopyMemory theLong, ARGB, 4
End Sub

Public Function GetWord(ByVal strVal As String) As Long

    Dim Lo As Long

    Dim Hi As Long
    
    Lo = AscB(MidB(strVal, 1, 1))
    Hi = AscB(MidB(strVal, 2, 1))
    
    GetWord = (Hi * 256) + Lo
End Function

Public Function GetDWord(ByVal strVal As String) As Double

    Dim LOWORD As Single

    Dim HiWord As Single

    If LenB(strVal) <> 4 Then
        GetDWord = 0

        Exit Function

    End If
    
    LOWORD = GetWord(MidB(strVal, 1, 2))
    HiWord = GetWord(MidB(strVal, 3, 2))
    GetDWord = (HiWord * 65536) + LOWORD
End Function
