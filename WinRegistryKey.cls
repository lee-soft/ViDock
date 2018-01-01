VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinRegistryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : WinRegistryKey
'    Project    : ViDock
'
'    Description: A registry class for manipulating the registry through pure
'                 WinAPI
'
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

Private Declare Function RegEnumKey _
                Lib "advapi32.dll" _
                Alias "RegEnumKeyA" (ByVal hKey As Long, _
                                     ByVal dwIndex As Long, _
                                     ByVal lpName As String, _
                                     ByVal cbName As Long) As Long

Private Declare Function RegQueryValue _
                Lib "advapi32.dll" _
                Alias "RegQueryValueA" (ByVal hKey As Long, _
                                        ByVal lpSubKey As String, _
                                        ByVal lpValue As String, _
                                        lpcbValue As Long) As Long

Private m_Handle    As Long

Private m_SubKeys   As Collection

Private m_Path      As String

Private m_Name      As String

Private m_LastError As Long

Private m_rootKey   As EROOTKEY

Friend Property Let RootKeyType(newRootKeyType As EROOTKEY)
    m_rootKey = newRootKeyType
    pEnumerateSubKeys
    
End Property

Public Function GetValueAsString(Optional KeyName As String = vbNullString)

    On Error GoTo Handler

    Dim length       As Long

    Dim subkey_value As String

    Dim sKeyType     As EREGTYPE

    m_LastError = 0

    If KeyName <> vbNullString Then
        ' Set up buffer for data to be returned in.
        ' Adjust next value for larger buffers.
        length = 255
        subkey_value = Space$(length)
    
        ' Read key
        If RegQueryValueEx(m_Handle, KeyName, ByVal 0&, sKeyType, StrPtr(subkey_value), length) <> ERROR_SUCCESS Then
            pErrorHelper -1, "Key Name: " & KeyName
        Else

            ' Remove the trailing null character.
            If sKeyType = REG_SZ Or sKeyType = REG_BINARY Then
                
                subkey_value = MidB(subkey_value, 1, length - 2)
            Else
                pErrorHelper -4, "Key Name: " & KeyName
            End If
        End If

    Else
        ' Get the subkey's value.
        length = 256
        subkey_value = Space$(length)
        
        If RegQueryValue(m_Handle, vbNullString, subkey_value, length) <> ERROR_SUCCESS Then
            pErrorHelper -1, ""
        Else

            ' Remove the trailing null character.
            If length > 0 Then
                subkey_value = Left$(subkey_value, length - 1)
            End If
        End If
    End If

    GetValueAsString = subkey_value
    
    Exit Function

Handler:
    LogError 1001, Err.Description, "WinRegistryKey"
End Function

Public Property Let Path(ByVal newPath As String)

    ' Open the key.
    If RegOpenKeyEx(m_rootKey, newPath, 0&, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE, m_Handle) <> ERROR_SUCCESS Then

        Exit Property

    End If
    
    m_Path = newPath
    pEnumerateSubKeys
    
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Get SubKeys() As Collection
    Set SubKeys = m_SubKeys
End Property

Public Property Get SubKey(index As Long) As WinRegistryKey
    Set SubKey = m_SubKeys(index)
    
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Friend Property Let Handle(ByVal newHandle As Long)
    m_Handle = newHandle
    
    'pEnumerateSubKeys
End Property

Friend Property Let Name(ByVal newName As String)
    m_Name = newName
End Property

Private Function pEnumerateSubKeys()

    On Error Resume Next

    If m_rootKey = 0 Then
        m_LastError = -2

        Exit Function

    End If
    
    If m_Handle = 0 Then
        m_LastError = -3

        Exit Function

    End If
    
    m_LastError = 0

    Dim subkey_num  As Integer

    Dim subkey_name As String

    Dim length      As Long

    Dim hKey        As Long

    Dim newSubKey   As WinRegistryKey
    
    Set m_SubKeys = New Collection

    ' Enumerate the subkeys.
    subkey_num = 0
    hKey = m_Handle
    
    Do
        ' Enumerate subkeys until we get an error.
        length = 256
        subkey_name = Space$(length)

        If RegEnumKey(hKey, subkey_num, subkey_name, length) <> ERROR_SUCCESS Then Exit Do
                
        subkey_num = subkey_num + 1
        subkey_name = Left$(subkey_name, InStr(subkey_name, Chr$(0)) - 1)
        
        Set newSubKey = New WinRegistryKey
        newSubKey.Name = subkey_name
        newSubKey.RootKeyType = m_rootKey
        
        m_SubKeys.Add newSubKey
        newSubKey.Path = m_Path & "\" & subkey_name
        
    Loop

End Function

Private Sub Class_Terminate()

    If (m_Handle <> 0) Then
        RegCloseKey m_Handle
    End If

End Sub

Private Sub pErrorHelper(newErrorCode As Long, Optional additionalInfo As String)
    m_LastError = newErrorCode
    
    If m_LastError = -1 Then
        'Debug.Print "Error querying key: " & m_Path & vbCrLf & _
         additionalInfo
                    
    ElseIf m_LastError = -2 Then
    
    ElseIf m_LastError = -3 Then
        
    ElseIf m_LastError = -4 Then
        Debug.Print "Type was not convertable to a string: " & m_Path & vbCrLf & additionalInfo
    End If

End Sub

Public Function GetLastError()
    GetLastError = m_LastError
End Function
