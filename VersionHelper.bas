Attribute VB_Name = "VersionHelper"
'--------------------------------------------------------------------------------
'    Component  : VersionHelper
'    Project    : ViDock
'
'    Description: Determines the file version
'
'--------------------------------------------------------------------------------
Private Declare Function lstrcpy Lib "kernel32" _
         Alias "lstrcpyA" _
         (ByVal lpString1 As String, _
         ByVal lpString2 As Long) As Long
         
Private Declare Function VerQueryValue Lib "Version.dll" _
         Alias "VerQueryValueA" _
         (pBlock As Any, _
         ByVal lpSubBlock As String, _
         lplpBuffer As Any, _
         puLen As Long) As Long

Public Function GetEXEProductTitle(ByVal szExePath As String) As String

Dim lVerPointer As Long
Dim sBuffer()  As Byte
Dim rc As Long
Dim lBufferLen As Long, lDummy As Long
Dim bytebuffer(255) As Byte
Dim Lang_Charset_String As String
Dim HexNumber As Long
Dim Buffer As String
Dim strTemp As String

    Buffer = String(255, 0)
    
    '*** Get size ****
    lBufferLen = GetFileVersionInfoSize(szExePath, lDummy)

    If lBufferLen < 1 Then
       Exit Function
    End If
    
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(szExePath, _
                            0&, _
                            lBufferLen, _
                            sBuffer(0))
    If rc = 0 Then
       Exit Function
    End If
    
    rc = VerQueryValue(sBuffer(0), _
                           "\VarFileInfo\Translation", _
                           lVerPointer, _
                           lBufferLen)

    If rc = 0 Then
       Exit Function
    End If
    
    'lVerPointer is a pointer to four 4 bytes of Hex number,
    'first two bytes are language id, and last two bytes are code
    'page. However, Lang_Charset_String needs a  string of
    '4 hex digits, the first two characters correspond to the
    'language id and last two the last two character correspond
    'to the code page id.
    
    RtlMoveMemory bytebuffer(0), ByVal lVerPointer, ByVal lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + _
         bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
         Lang_Charset_String = Hex(HexNumber)
         
    'now we change the order of the language id and code page
    'and convert it into a string representation.
    'For example, it may look like 040904E4
    'Or to pull it all apart:
    '04------        = SUBLANG_ENGLISH_USA
    '--09----        = LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows:Multilingual
    Do While Len(Lang_Charset_String) < 8
       Lang_Charset_String = "0" & Lang_Charset_String
    Loop
    
    Buffer = String(255, 0)
    strTemp = "\StringFileInfo\" & Lang_Charset_String _
    & "\FileDescription"
    
    rc = VerQueryValue(sBuffer(0), strTemp, _
    lVerPointer, lBufferLen)

    If rc = 0 Then
        Exit Function
    End If

    lstrcpy Buffer, lVerPointer
    Buffer = Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1)
    GetEXEProductTitle = Buffer
End Function



