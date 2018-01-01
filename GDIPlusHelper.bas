Attribute VB_Name = "GDIPlusHelper"
Option Explicit



Public Function CreateFontFamily(szFontName As String) As GDIPFontFamily

Dim newFontFamily As New GDIPFontFamily
    newFontFamily.Constructor szFontName

    Set CreateFontFamily = newFontFamily
End Function

