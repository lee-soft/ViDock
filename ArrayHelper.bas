Attribute VB_Name = "ArrayHelper"
Option Explicit

Public Function UniqueValues(ByRef heyStack)

    Dim m As Long

    Dim n As Long

    Dim newArray()

    If Not IsArrayInitialized(heyStack) Then

        Exit Function

    End If
    
    For m = LBound(heyStack) To UBound(heyStack)

        If Not In_Array(newArray, heyStack(m)) Then
            ReDim Preserve newArray(n)
            
            newArray(n) = heyStack(m)
            n = n + 1
        End If

    Next
    
    UniqueValues = newArray

End Function

Public Function In_Array(ByRef a, ByRef sValue) As Boolean

    Dim m As Long

    If Not IsArrayInitialized(a) Then
        In_Array = False

        Exit Function

    End If
    
    For m = LBound(a) To UBound(a)

        If (a(m) = sValue) Then
            In_Array = True

            Exit Function

        End If

    Next

End Function

Public Function ConcatArray(ByRef a, ByRef b)

    Dim n As Long, m As Long

    Dim c()

    n = 0
    
    If IsArrayInitialized(a) Then

        For m = LBound(a) To UBound(a)
            ReDim Preserve c(n)
            
            c(n) = a(m)
            n = n + 1
        Next

    End If
    
    If IsArrayInitialized(b) Then

        For m = LBound(b) To UBound(b)
            ReDim Preserve c(n)
            
            c(n) = b(m)
            n = n + 1
        Next

    End If
    
    ConcatArray = c

End Function

Public Function IsArrayInitialized(myArray) As Boolean

    Dim mySize As Long

    On Error Resume Next

    mySize = UBound(myArray) ' In this instance the error number is set as myArray has a size of -1!

    If (Err.Number <> 0) Then
        IsArrayInitialized = False
    Else

        If mySize > -1 Then
            IsArrayInitialized = True
        End If
    End If

End Function

Public Function sizeOf(srcArray) As Long

    On Error GoTo Handler
    
    sizeOf = UBound(srcArray)

    Exit Function

Handler:
    sizeOf = 0
    
End Function
