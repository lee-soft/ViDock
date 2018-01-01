Attribute VB_Name = "MathHelper"
'--------------------------------------------------------------------------------
'    Component  : MathHelper
'    Project    : ViDock
'
'    Description: Contains Math helper functions
'
'--------------------------------------------------------------------------------
Option Explicit

Public Function RoundIt(ByVal lngSrcNumber As Integer, ByVal lngByNumber As Integer)

    'Round(12, 5)
   
    Dim lngModResult As Long

    lngModResult = (lngSrcNumber Mod lngByNumber)
   
    If lngModResult >= lngByNumber Then
        RoundIt = CLng(SymArith(lngSrcNumber / lngByNumber, 0) * lngByNumber + 1)
    Else
        RoundIt = CLng(SymArith(lngSrcNumber / lngByNumber, 0) * lngByNumber)
    End If

End Function

Public Function SymArith(ByVal X As Double, _
                         Optional ByVal DecimalPlaces As Double = 1) As Double

    SymArith = Fix(X * (10 ^ DecimalPlaces) + 0.5 * Sgn(X)) / (10 ^ DecimalPlaces)
End Function


Public Function Floor(ByVal Number) As Long
    
    Floor = Fix(Number)
    
    If Number >= 0 Then
            
        If Number = Int(Number) Then
            
            Floor = Number
        
        Else
        
            Floor = Int(Number)
            
        End If
        
    ElseIf Number < 0 Then
    
        Floor = Int(Number) - 1
        
    End If
 
End Function

Public Function Ceiling(ByVal Number) As Long
 
    If Number >= 0 Then
            
        If Number = Int(Number) Then
            
            Ceiling = Number
        
        Else
        
            Ceiling = Int(Number) + 1
            
        End If
        
    ElseIf Number < 0 Then
    
        Ceiling = Int(Number)
        
    End If
 
End Function

Public Function Sqrt(X As Double) As Double
    Sqrt = X ^ (1 / 2)
End Function

