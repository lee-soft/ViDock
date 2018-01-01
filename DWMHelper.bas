Attribute VB_Name = "DWMHelper"
'Module Name    Desktop Window Manager Declare for Visual Basic 6
'Author                         §®§Ñ§Ô§ã§ß
'Version                        0.0.1
'You are free to use this module.
'But you should keep the copyright information.

Option Explicit

Public Declare Function CopyMemory2 _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (ByVal pDest As Any, _
                                      pSource As Any, _
                                      dwLength As Long) As Long

Private Declare Function CopyFromRectToPointer _
                Lib "user32" _
                Alias "CopyRect" (ByVal lpDestRect As Long, _
                                  lpSourceRect As gdiplus.RECTL) As Long


Public Const DWM_BB_ENABLE = &H1

Public Const DWM_BB_BLURREGION = &H2

Public Const DWM_BB_TRANSITIONONMAXIMIZED = &H4

Public Type DWM_BLURBEHIND

    dwFlags As Long
    fEnable As Long
    hRgnBlur As Long
    fTransitionOnMaximized As Long

End Type

Public Type RECT

    Left As Long
    Bottom As Long
    Right As Long
    Top As Long

End Type

Public Enum DWMWINDOWATTRIBUTE

    DWMWA_NCRENDERING_ENABLED = 1                       '[get] Is non-client rendering enabled/disabled
    DWMWA_NCRENDERING_POLICY = 2                          '[set] Non-client rendering policy
    DWMWA_TRANSITIONS_FORCEDISABLED = 3              '[set] Potentially enable/forcibly disable transitions
    DWMWA_ALLOW_NCPAINT = 4                                   '[set] Allow contents rendered in the non-client area to be visible on the DWM-drawn frame.
    DWMWA_CAPTION_BUTTON_BOUNDS = 5                  '[get] Bounds of the caption button area in window-relative space.
    DWMWA_NONCLIENT_RTL_LAYOUT = 6                      '[set] Is non-client content RTL mirrored
    DWMWA_FORCE_ICONIC_REPRESENTATION = 7         '[set] Force this window to display iconic thumbnails.
    DWMWA_FLIP3D_POLICY = 8                                      '[set] Designates how Flip3D will treat the window.
    DWMWA_EXTENDED_FRAME_BOUNDS = 8                 '[get] Gets the extended frame bounds rectangle in screen space
    DWMWA_LAST

End Enum

Public Enum DWMNCRENDERINGPOLICY

    DWMNCRP_USEWINDOWSTYLE = 0    'Enable/disable non-client rendering based on window style
    DWMNCRP_DISABLED = 1                   'Disabled non-client rendering; window style is ignored
    DWMNCRP_ENABLED = 2                    'Enabled non-client rendering; window style is ignored
    DWMNCRP_LAST = 3

End Enum

Public Enum DWMFLIP3DWINDOWPOLICY

    DWMFLIP3D_DEFAULT = 0                     'Hide or include the window in Flip3D based on window style and visibility.
    DWMFLIP3D_EXCLUDEBELOW = 1          'Display the window under Flip3D and disabled.
    DWMFLIP3D_EXCLUDEABOVE = 2           'Display the window above Flip3D and enabled.
    DWMFLIP3D_LAST = 3

End Enum

'typedef HANDLE HTHUMBNAIL;
'typedef HTHUMBNAIL* PHTHUMBNAIL;

Public Const DWM_TNP_RECTDESTINATION = &H1

Public Const DWM_TNP_RECTSOURCE = &H2

Public Const DWM_TNP_OPACITY = &H4

Public Const DWM_TNP_VISIBLE = &H8

Public Const DWM_TNP_SOURCECLIENTAREAONLY = &H10

Public Type DWM_THUMBNAIL_PROPERTIES

    dwFlags As Long
    rcDestination As RECT 'RECT rcDestination;
    rcSource As RECT     'RECT rcSource;
    opacity As Byte         'BYTE opacity
    fVisible As Long
    fSourceClientAreaOnly As Long

End Type

'typedef ULONGLONG DWM_FRAME_COUNT;
'typedef ULONGLONG QPC_TIME;

Public Type UNSIGNED_RATIO

    uiNumerator As Long 'UInt32
    uiDenominator As Long 'UInt32

End Type

Public Type DWM_TIMING_INFO

    cbSize As Long '      UINT32 cbSize

    'Data on DWM composition overall
    
    'Monitor refresh rate
    rateRefresh As UNSIGNED_RATIO    'UNSIGNED_RATIO  rateRefresh;

    ' Actual period
    qpcRefreshPeriod As Double       'double     qpcRefreshPeriod;

    ' composition rate
    rateCompose As UNSIGNED_RATIO

    ' QPC time at a VSync interupt
    qpcVBlank As Double

    ' DWM refresh count of the last vsync
    ' DWM refresh count is a 64bit number where zero is
    ' the first refresh the DWM woke up to process
    cRefresh As Double 'double cRefresh;

    ' DX refresh count at the last Vsync Interupt
    ' DX refresh count is a 32bit number with zero
    ' being the first refresh after the card was initialized
    ' DX increments a counter when ever a VSync ISR is processed
    ' It is possible for DX to miss VSyncs
    '
    ' There is not a fixed mapping between DX and DWM refresh counts
    ' because the DX will rollover and may miss VSync interupts
    cDXRefresh As Long 'UINT cDXRefresh;

    '// QPC time at a compose time.
    qpcCompose As Double 'double        qpcCompose;

    ' Frame number that was composed at qpcCompose
    cFrame As Double 'double cFrame;

    'The present number DX uses to identify renderer frames
    cDXPresent As Long  'UINT            cDXPresent;

    ' Refresh count of the frame that was composed at qpcCompose
    cRefreshFrame As Double 'double cRefreshFrame;

    ' DWM frame number that was last submitted
    cFrameSubmitted As Double 'double cFrameSubmitted;

    ' DX Present number that was last submitted
    cDXPresentSubmitted As Long 'UINT cDXPresentSubmitted;

    ' DWM frame number that was last confirmed presented
    cFrameConfirmed As Double 'double cFrameConfirmed;

    ' DX Present number that was last confirmed presented
    cDXPresentConfirmed As Long 'UINT cDXPresentConfirmed;

    ' The target refresh count of the last
    ' frame confirmed completed by the GPU
    cRefreshConfirmed As Double 'double cRefreshConfirmed;

    ' DX refresh count when the frame was confirmed presented
    cDXRefreshConfirmed As Long 'UINT cDXRefreshConfirmed;

    ' Number of frames the DWM presented late
    ' AKA Glitches
    cFramesLate As Double 'double          cFramesLate;
    
    ' the number of composition frames that
    ' have been issued but not confirmed completed
    cFramesOutstanding As Long 'UINT          cFramesOutstanding;

    ' Following fields are only relavent when an HWND is specified
    ' Display frame

    ' Last frame displayed
    cFrameDisplayed As Double 'double cFrameDisplayed;

    ' QPC time of the composition pass when the frame was displayed
    qpcFrameDisplayed As Double 'double        qpcFrameDisplayed;

    ' Count of the VSync when the frame should have become visible
    cRefreshFrameDisplayed As Double 'double cRefreshFrameDisplayed;

    ' Complete frames: DX has notified the DWM that the frame is done rendering

    ' ID of the the last frame marked complete (starts at 0)
    cFrameComplete As Double 'double cFrameComplete;

    ' QPC time when the last frame was marked complete
    qpcFrameComplete As Double 'double        qpcFrameComplete;

    ' Pending frames:
    ' The application has been submitted to DX but not completed by the GPU
 
    ' ID of the the last frame marked pending (starts at 0)
    cFramePending As Double 'double cFramePending;

    ' QPC time when the last frame was marked pending
    qpcFramePending As Double 'double        qpcFramePending;

    ' number of unique frames displayed
    cFramesDisplayed As Double 'double cFramesDisplayed;

    ' number of new completed frames that have been received
    cFramesComplete As Double 'double cFramesComplete;

    ' number of new frames submitted to DX but not yet complete
    cFramesPending As Double 'double cFramesPending;

    ' number of frames available but not displayed, used or dropped
    cFramesAvailable As Double 'double cFramesAvailable;

    ' number of rendered frames that were never
    ' displayed because composition occured too late
    cFramesDropped As Double 'double cFramesDropped;
    
    ' number of times an old frame was composed
    ' when a new frame should have been used
    ' but was not available
    cFramesMissed As Double 'double cFramesMissed;
    
    ' the refresh at which the next frame is
    ' scheduled to be displayed
    cRefreshNextDisplayed As Double 'double cRefreshNextDisplayed;

    ' the refresh at which the next DX present is
    ' scheduled to be displayed
    cRefreshNextPresented As Double 'double cRefreshNextPresented;

    ' The total number of refreshes worth of content
    ' for this HWND that have been displayed by the DWM
    ' since DwmSetPresentParameters was called
    cRefreshesDisplayed As Double  'double cRefreshesDisplayed;
    
    ' The total number of refreshes worth of content
    ' that have been presented by the application
    ' since DwmSetPresentParameters was called
    cRefreshesPresented As Double 'double cRefreshesPresented;

    ' The actual refresh # when content for this
    ' window started to be displayed
    ' it may be different than that requested
    ' DwmSetPresentParameters
    cRefreshStarted As Double 'double cRefreshStarted;

    ' Total number of pixels DX redirected
    ' to the DWM.
    ' If Queueing is used the full buffer
    ' is transfered on each present.
    ' If not queuing it is possible only
    ' a dirty region is updated
    cPixelsReceived As Double  'ULONGLONG  cPixelsReceived;

    ' Total number of pixels drawn.
    ' Does not take into account if
    ' if the window is only partial drawn
    ' do to clipping or dirty rect management
    cPixelsDrawn As Double 'ULONGLONG  cPixelsDrawn;

    ' The number of buffers in the flipchain
    ' that are empty.   An application can
    ' present that number of times and guarantee
    ' it won't be blocked waiting for a buffer to
    ' become empty to present to
    cBuffersEmpty As Double 'double      cBuffersEmpty;

End Type

Public Enum DWM_SOURCE_FRAME_SAMPLING

    '// Use the first source frame that
    '// includes the first refresh of the output frame
    DWM_SOURCE_FRAME_SAMPLING_POINT = &H0

    '// use the source frame that includes the most
    '// refreshes of out the output frame
    '// in case of multiple source frames with the
    '// same coverage the last will be used
    DWM_SOURCE_FRAME_SAMPLING_COVERAGE = &H1

    '// Sentinel value
    DWM_SOURCE_FRAME_SAMPLING_LAST = &H2

End Enum

Public Const c_DwmMaxQueuedBuffers As Long = 8

Public Const c_DwmMaxMonitors      As Long = 16

Public Const c_DwmMaxAdapters      As Long = 16

Public Type DWM_PRESENT_PARAMETERS

    vbSize As Long ' UINT32          cbSize;
    fQueue As Long 'BOOL            fQueue;
    cRefreshStart As Double 'double cRefreshStart;
    cBuffer As Long 'UINT            cBuffer;
    fUseSourceRate As Long 'BOOL            fUseSourceRate;
    rateSource As UNSIGNED_RATIO 'UNSIGNED_RATIO  rateSource;
    cRefreshesPerFrame As Long 'UINT            cRefreshesPerFrame;
    eSampling As DWM_SOURCE_FRAME_SAMPLING 'DWM_SOURCE_FRAME_SAMPLING  eSampling;

End Type

Public Const DWM_FRAME_DURATION_DEFAULT = -1

Public Declare Function DwmDefWindowProc _
               Lib "dwmapi.dll" (hWnd As Long, _
                                 msg As Long, _
                                 wParam As Long, _
                                 lParam As Long, _
                                 lResult As Long) As Long

Public Declare Function DwmEnableBlurBehindWindow _
               Lib "dwmapi.dll" (ByVal hWnd As Long, _
                                 ByRef pBlurBehind As DWM_BLURBEHIND) As Long

Public Const DWM_EC_DISABLECOMPOSITION = 0

Public Const DWM_EC_ENABLECOMPOSITION = 1

Public Declare Function DwmEnableComposition Lib "dwmapi.dll" (uCompositionAction As Long)

Public Declare Function DwmEnableMMCSS _
               Lib "dwmapi" (fEnableMMCSS As Long) 'BOOL fEnableMMCSS

Public Declare Function DwmExtendFrameIntoClientArea _
               Lib "dwmapi" (ByVal hWnd As Long, _
                             pMarInset As gdiplus.RECTL) As Long

Public Type Margins

    m_Left As Long
    m_Right As Long
    m_Top As Long
    m_Bottom As Long

End Type

Public Declare Function DwmGetColorizationColor _
               Lib "dwmapi.dll" (pcrColorization As Long, _
                                 pfOpaqueBlend As Long)

Public Declare Function DwmGetCompositionTimingInfo _
               Lib "dwmapi.dll" (hWnd As Long, _
                                 pTimingInfo As DWM_TIMING_INFO)

Public Declare Function DwmGetWindowAttribute _
               Lib "dwmapi.dll" (hWnd As Long, _
                                 dwAttribute As Long, _
                                 pvAttribute As Long, _
                                 cbAttribute As Long)
'    __out_bcount(cbAttribute) PVOID pvAttribute,

Public Declare Function DwmIsCompositionEnabled _
               Lib "dwmapi.dll" (pfEnabled As Long) As Long

Public Declare Function DwmModifyPreviousDxFrameDuration _
               Lib "dwmapi.dll" (hWnd As Long, _
                                 cRefreshes As Long, _
                                 fRelative As Long)

Public Declare Function DwmQueryThumbnailSourceSize _
               Lib "dwmapi.dll" (hThumbnail As Long, _
                                 psize As Long) '    HTHUMBNAIL hThumbnail,

Public Type Size

    cx As Long
    cy As Long

End Type

Public Declare Function DwmRegisterThumbnail _
               Lib "dwmapi.dll" (hwndDestination As Long, _
                                 hwndSource As Long, _
                                 phThumbnailId As Long)

Public Declare Function DwmSetDxFrameDuration _
               Lib "dwmapi.dll" (hWnd As Long, _
                                 cRefreshes As Long)

Public Declare Function DwmSetPresentParameters _
               Lib "dwmapi.dll" (hWnd As Long, _
                                 pPresentParams As Long) 'pPresentParams =lp DWM_PRESENT_PARAMETERS)

Public Declare Function DwmSetWindowAttribute _
               Lib "dwmapi.dll" (hWnd As Long, _
                                 dwAttribute As Long, _
                                 pvAttribute As Long, _
                                 cbAttribute As Long)

Public Declare Function DwmUnregisterThumbnail Lib "dwmapi.dll" (hThumbnailId As Long)

Public Declare Function DwmUpdateThumbnailProperties _
               Lib "dwmapi.dll" (hThumbnailId As Long, _
                                 ptnProperties As Long) 'ptnProperties = lp DWM_THUMBNAIL_PROPERTIES

Public Declare Function DwmAttachMilContent Lib "dwmapi.dll" (hWnd As Long)

Public Declare Function DwmDetachMilContent Lib "dwmapi.dll" (hWnd As Long)

Public Declare Function DwmFlush Lib "dwmapi.dll" ()

Public Type MIL_MATRIX3X2D

    S_11 As Double
    S_12 As Double
    S_21 As Double
    S_22 As Double
    Dx As Double
    Dy As Double

End Type

Public Declare Function DwmGetGraphicsStreamTransformHint _
               Lib "dwmapi.dll" (uIndex As Long, _
                                 pTransform As Long) 'pTransform = lp MIL_MATRIX3X2D

Public Declare Function DwmGetGraphicsStreamClien _
               Lib "dwmapi.dll" (uIndex As Long, _
                                 pClientUuid As Long) 'pClientUuid= lp UUID

Public Declare Function DwmGetTransportAttributes _
               Lib "dwmapi.dll" (pfIsRemoting As Long, _
                                 pfIsConnected As Long, _
                                 pDwGeneration As Long)

Public Function GetRectPointer(ByRef theRect As gdiplus.RECTL) As Long

    Dim pRect As Long

    CopyFromRectToPointer pRect, theRect
    GetRectPointer = pRect
End Function

