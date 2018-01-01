VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_hRegion    As Long

Private m_lastResult As Long

Public Function GetBounds(ByRef theGraphics As GDIPGraphics) As RECTF

    Dim theBounds As RECTF

    SetStatus GdipGetRegionBounds(m_hRegion, thegrahpics, theBounds)
    GetBounds = theBounds
End Function

Public Sub Constructor(ByRef theRect As RECTF)
    Dispose
    GdipCreateRegionRect theRect, m_hRegion
End Sub

Public Property Get Value() As Long
    Value = m_hRegion
End Property

Private Sub Class_Initialize()
    Dispose
    GdipCreateRegion m_hRegion
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

Private Sub Dispose()

    If m_hRegion <> 0 Then
        GdipDeleteRegion m_hRegion
    End If

End Sub

Private Function SetStatus(ByVal status As GpStatus) As GpStatus
    m_lastResult = SetStatusHelper(status)
End Function
