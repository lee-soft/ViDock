Attribute VB_Name = "IconHelper"
Option Explicit

Private Const IID_IImageList  As String = "{46EB5926-582E-4017-9FDF-E8998DAA0950}"

Private Const IID_IImageList2 As String = "{192B9D83-50FC-457B-90A0-2B82A8B5DAE1}"

Private Const E_INVALIDARG    As Long = &H80070057

Private Const ILD_NORMAL      As Long = 0

Private Const ILD_TRANSPARENT = &H1 'display transparent

Private Const SHGFI_DISPLAYNAME = &H200

Private Const SHGFI_EXETYPE = &H2000

Private Const SHGFI_SYSICONINDEX = &H4000 'system icon index

Private Const SHGFI_LARGEICON = &H0 'large icon

Private Const SHGFI_SMALLICON = &H1 'small icon

Private Const SHGFI_EXTRALARGE = &H2

Private Const SHGFI_SHELLICONSIZE = &H4

Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Enum SHIL_FLAG

    SHIL_LARGE = &H0      '   The image size is normally 32x32 pixels. However, if the Use large icons option is selected from the Effects section of the Appearance tab in Display Properties, the image is 48x48 pixels.
    SHIL_SMALL = &H1      '   These images are the Shell standard small icon size of 16x16, but the size can be customized by the user.
    SHIL_EXTRALARGE = &H2 '   These images are the Shell standard extra-large icon size. This is typically 48x48, but the size can be customized by the user.
    SHIL_SYSSMALL = &H3   '   These images are the size specified by GetSystemMetrics called with SM_CXSMICON and GetSystemMetrics called with SM_CYSMICON.
    SHIL_JUMBO = &H4      '   Windows Vista and later. The image is normally 256x256 pixels.

End Enum

Private Function DrawIconToHDC(aFile As String, theHDC As Long)

    Dim aImgList As Long

    Dim SFI      As SHFILEINFO

    SHGetFileInfo aFile, FILE_ATTRIBUTE_NORMAL, SFI, Len(SFI), SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_DISPLAYNAME
                              
    aImgList = GetImageListSH(SHIL_JUMBO)
    
    ImageList_Draw aImgList, SFI.iIcon, theHDC, 0, 0, ILD_NORMAL

End Function

Public Function IconIs48(aFile As String) As Boolean

    Dim newHDC As New cMemDC

    Dim pixelA As Long

    Dim pixelB As Long
    
    newHDC.Height = 256
    newHDC.Width = 256
    
    DrawIconToHDC aFile, newHDC.hdc
    
    IconIs48 = True
    
    For pixelA = 48 To 255
        For pixelB = 48 To 255

            If GetPixel(newHDC.hdc, pixelA, pixelB) <> 0 Then
                IconIs48 = False

                Exit Function

            End If

        Next
    Next

End Function

Private Function GetImageListSH(shFlag As SHIL_FLAG) As Long

    Dim lResult      As Long

    Dim Guid(0 To 3) As Long

    Dim himl         As IUnknown

    If Not IIDFromString(StrPtr(IID_IImageList), Guid(0)) = 0 Then

        Exit Function

    End If
    
    lResult = SHGetImageListXP(CLng(shFlag), Guid(0), ByVal VarPtr(himl))
    GetImageListSH = ObjPtr(himl)
End Function

Public Function GetIconFromHwnd(hWnd As Long) As Long
    Call SendMessageTimeout(hWnd, WM_GETICON, ICON_BIG, 0, 0, 100, GetIconFromHwnd)

    If Not CBool(GetIconFromHwnd) Then GetIconFromHwnd = GetClassLong(hWnd, GCL_HICON)
    If Not CBool(GetIconFromHwnd) Then Call SendMessageTimeout(hWnd, WM_GETICON, 1, 0, 0, 100, GetIconFromHwnd)
    If Not CBool(GetIconFromHwnd) Then GetIconFromHwnd = GetClassLong(hWnd, GCL_HICON)
    If Not CBool(GetIconFromHwnd) Then Call SendMessageTimeout(hWnd, WM_QUERYDRAGICON, 0, 0, 0, 100, GetIconFromHwnd)
End Function

Public Function GetSmallApplicationIcon(strExePath As String) As Long

    Dim shinfo     As SHFILEINFO

    Dim hImgSmall  As Long

    Dim win64Token As Win64FSToken

    If Is64bit Then
        If (InStr(LCase(strExePath), LCase(Environ("windir"))) > 0) Then
            Set win64Token = New Win64FSToken
        End If
    End If

    'get the system icon associated with that file
    hImgSmall = WinAPIHelper.SHGetFileInfo(strExePath, 0&, shinfo, Len(shinfo), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_DISPLAYNAME)

    GetSmallApplicationIcon = shinfo.hIcon

    If Not win64Token Is Nothing Then
        win64Token.EnableFS
    End If

End Function

Function CreateSmallAlphaIcon(szPath As String) As AlphaIcon

    Dim SmallIcon    As Long

    Dim FileInfo     As SHFILEINFO

    Dim newAlphaIcon As AlphaIcon

    SmallIcon = SHGetFileInfo(szPath, 0&, FileInfo, Len(FileInfo), SHGFI_SMALLICON Or SHGFI_ICON)
    
    Set newAlphaIcon = New AlphaIcon
    newAlphaIcon.CreateFromHICON FileInfo.hIcon
    
    DestroyIcon FileInfo.hIcon
    Set CreateSmallAlphaIcon = newAlphaIcon

End Function

Public Function GetExtraLargeApplicationIcon(szPath As String) As Long

    Dim aImgList As Long

    Dim SFI      As SHFILEINFO

    Dim hIcon    As Long

    SHGetFileInfo szPath, FILE_ATTRIBUTE_NORMAL, SFI, Len(SFI), SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_DISPLAYNAME

    If Not IconIs48(szPath) Then
        aImgList = GetImageListSH(SHIL_JUMBO)
    Else
        aImgList = GetImageListSH(SHIL_EXTRALARGE)
    End If
        
    hIcon = ImageList_GetIcon(aImgList, SFI.iIcon, ILD_NORMAL)
    GetExtraLargeApplicationIcon = hIcon
End Function

Public Function GetApplicationIcon(strExePath As String) As Long

    Dim shinfo     As SHFILEINFO

    Dim hImgSmall  As Long

    Dim win64Token As Win64FSToken

    If Is64bit Then
        If (InStr(LCase(strExePath), LCase(Environ("windir"))) > 0) Then
            Set win64Token = New Win64FSToken
        End If
    End If

    'get the system icon associated with that file
    hImgSmall = WinAPIHelper.SHGetFileInfo(strExePath, 0&, shinfo, Len(shinfo), SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_DISPLAYNAME)

    GetApplicationIcon = shinfo.hIcon

    If Not win64Token Is Nothing Then
        win64Token.EnableFS
    End If

End Function
