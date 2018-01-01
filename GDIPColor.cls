VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Red   As Byte

Private m_Green As Byte

Private m_Blue  As Byte

Private m_Alpha As Byte

Public Event onChanged()

Public Property Get Alpha() As Byte
    Alpha = m_Alpha
End Property

Public Property Let Alpha(ByVal v_Alpha As Byte)
    m_Alpha = v_Alpha
End Property

Public Property Get Red() As Byte
    Red = m_Red
End Property

Public Property Let Red(ByVal v_Red As Byte)
    m_Red = v_Red
End Property

Public Property Get Blue() As Byte
    Blue = m_Blue
End Property

Public Property Let Blue(ByVal v_Blue As Byte)
    m_Blue = v_Blue
End Property

Public Property Get Green() As Byte
    Green = m_Green
End Property

Public Property Let Green(ByVal v_Green As Byte)
    m_Green = v_Green
End Property

Public Property Let Value(ByVal vData As Long)
    'Debug.Print "Colour::Value::Let: " & vData

    Long2RGB vData
    RaiseEvent onChanged
End Property

Public Property Get Value() As Long
    Value = RGB(m_Red, m_Green, m_Blue)
End Property

Public Sub SetColourByHex(ByRef strHexColor As String)

    Dim HexColor As String

    Dim i        As Byte
    
    On Error Resume Next

    ' make sure the string is 6 characters l
    '     ong
    ' (it may have been given in &H###### fo
    '     rmat, we want ######)
    strHexColor = Strings.Right(strHexColor, 6)
    ' however, it may also have been given a
    '     s or #***** format, so add 0's in front

    For i = 1 To (6 - Len(strHexColor))
        HexColor = HexColor & "0"
    Next

    HexColor = HexColor & strHexColor
    ' convert each set of 2 characters into
    '     bytes, using vb's cbyte function
    m_Red = CByte("&H" & Strings.Right$(HexColor, 2))
    m_Green = CByte("&H" & Mid$(HexColor, 3, 2))
    m_Blue = CByte("&H" & Strings.Left$(HexColor, 2))
    m_Alpha = 255
    
    RaiseEvent onChanged
End Sub

Private Sub Long2RGB(LongColor As Long)

    On Error Resume Next

    ' convert to hex using vb's hex function
    '     , then use the hex2rgb function
    SetColourByHex (Hex(LongColor))
End Sub

Private Function Long2Hex(LongColor As Long) As String

    On Error Resume Next

    ' use vb's hex function
    Long2Hex = Hex(LongColor)
End Function

Private Sub Class_Initialize()
    SetColourByHex "#ffffff"
End Sub

