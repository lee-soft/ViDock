VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ======================================================================================
' Name:     cMemDC.cls
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     20 October 1999
'
' Requires: -
'
' Copyright © 1999 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Memory DC for flicker free drawing.
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================
'
' Declerations removed since they exist in the TLB

Option Explicit

Private m_hDC     As Long

Private m_hBmp    As Long

Private m_hBmpOld As Long

Private m_lWidth  As Long

Private m_lHeight As Long

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Let Width(ByVal Value As Long)

    If (Value > m_lWidth) Then
        m_lWidth = Value
        pCreate m_lWidth, m_lHeight
        SetBkMode m_hDC, 1
    End If

End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Let Height(ByVal Value As Long)

    If (Value > m_lHeight) Then
        m_lHeight = Value
        pCreate m_lWidth, m_lHeight
        SetBkMode m_hDC, 1
    End If

End Property

Public Property Get hdc() As Long
    hdc = m_hDC
End Property

Public Function CreateFromBitmap(strFile As String)
    'Creates a bitmap hdc
    'Returns the bitmap hdc
    pDestroy

    Dim lngBitMapDC As Long

    Dim vbPic       As IPictureDisp

    lngBitMapDC = CreateCompatibleDC(0)
    Set vbPic = LoadPicture(strFile)
    
    SelectObject lngBitMapDC, vbPic
    
    m_lHeight = vbPic.Height / 26.45454545455
    m_lWidth = vbPic.Width / 26.45454545455
    
    m_hDC = lngBitMapDC
    
End Function

Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)

    Dim lhDCC    As Long

    Dim initData As DEVMODE

    pDestroy
    lhDCC = CreateDC("DISPLAY", "", "", initData)

    If Not (lhDCC = 0) Then
        m_hDC = CreateCompatibleDC(lhDCC)

        If Not (m_hDC = 0) Then
            m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)

            If Not (m_hBmp = 0) Then
                m_hBmpOld = SelectObject(m_hDC, m_hBmp)

                If Not (m_hBmpOld = 0) Then
                    m_lWidth = Width
                    m_lHeight = Height
                    DeleteDC lhDCC

                    Exit Sub

                End If
            End If
        End If

        DeleteDC lhDCC
        pDestroy
    End If

End Sub

Private Sub pDestroy()

    If Not m_hBmpOld = 0 Then
        SelectObject m_hDC, m_hBmpOld
        m_hBmpOld = 0
    End If

    If Not m_hBmp = 0 Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If

    If Not m_hDC = 0 Then
        DeleteDC m_hDC
        m_hDC = 0
    End If

    m_lWidth = 0
    m_lHeight = 0
End Sub

Private Sub Class_Terminate()
    pDestroy
End Sub
