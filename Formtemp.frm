VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   737
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_graphics                As GDIPGraphics

Private m_bitmap                  As GDIPBitmap

Private m_pen                     As GDIPPen

Private m_black                   As Colour

Private m_layeredWindowProperties As LayerdWindowHandles

Private m_testImage               As GDIPImage

Private m_fontFamily              As GDIPFontFamily

Private m_font                    As GDIPFont

Private m_privateFontCollection   As GDIPlusWrapper.GDIPPrivateFC

Private m_privateFontFamily       As GDIPFontFamily

Private m_Y                       As Single

Private Sub Command1_Click()
    m_graphics.Clear vbWhite

    DrawThing "File"
    DrawThing "Edit"
    DrawThing "Format"
    DrawThing "View"
    DrawThing "Help"
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Load()
    Set m_fontFamily = CreateFontFamily("Segoe UI Semilight")
    
    Set m_privateFontCollection = New GDIPPrivateFC
    Set m_privateFontFamily = New GDIPFontFamily
    
    m_privateFontCollection.AddFontFile (App.Path & "\resources\font.ttf")
    m_privateFontCollection.AddFontFile (App.Path & "\resources\HelveticaNeue.ttf")
    
    m_privateFontFamily.Constructor2 "Helvetica Neue", m_privateFontCollection.NativeFontCollection
    
    Set m_font = New GDIPFont

    m_font.Constructor m_privateFontFamily, 15, FontStyleRegular
    'm_font.Constructor "Tahoma", 15, FontStyleRegular
    

    Set m_layeredWindowProperties = MakeLayerdWindow(Me)
    
    Set m_graphics = New GDIPGraphics
    Set m_testImage = New GDIPImage
    Set m_bitmap = New GDIPBitmap

    Set m_black = New Colour
    m_black.SetColourByHex "000000"
    
    Set m_pen = New GDIPPen
    m_pen.Constructor m_black, 1, 255
    m_testImage.FromFile App.Path & "\resources\bgfinder.png"
    'm_graphics.FromImage m_bitmap.Image
End Sub

Sub DrawThing(theText As String)

    Dim hDCOff As Long

    Dim hBMit  As Long

    Dim hSave  As Long

    Dim hCol   As Long

    'hCol = ColorARGB(0, 0, 0, 0)

    hDCOff = CreateCompatibleDC(0)
    m_bitmap.CreateFromSizeFormat Me.ScaleWidth, Me.ScaleHeight, PixelFormat32bppARGB
    hBMit = 0
    
    m_graphics.FromImage m_bitmap.Image
    'm_graphics.Clear2 ColorARGB(0, 255, 255, 255)

    m_graphics.TextRenderingHint = TextRenderingHintClearTypeGridFit
    m_graphics.CompositingMode = CompositingModeSourceOver
    m_graphics.CompositingQuality = CompositingQualityHighQuality
    m_graphics.PixelOffsetMode = PixelOffsetModeHighQuality

    m_graphics.DrawImage m_testImage, 0, 0, m_testImage.Width, m_testImage.Height

    'm_graphics.DrawRectangle m_pen, 5, 0, m_graphics.MeasureStringWidth(theText, AppDefaultFont), 17
    m_graphics.DrawString theText, m_font, GetBlackBrush, CreatePointF(5, CLng(m_Y))
    
    hBMit = m_bitmap.hBitmap(hCol)
    hSave = SelectObject(hDCOff, hBMit)
    
    m_layeredWindowProperties.Update Me.hWnd, 0, hDCOff
End Sub

Private Sub Timer1_Timer()
    Me.Move 0, 0
    'Me.DrawThing "Skype     Conversation       Call      Exit"
End Sub
