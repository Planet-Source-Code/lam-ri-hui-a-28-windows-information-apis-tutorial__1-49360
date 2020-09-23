VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "                Create Font Demo"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'used with fnWeight
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2
'used with SetBkMode
Const OPAQUE = 2
Const TRANSPARENT = 1

Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5
Const Message = "Hello !"

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Dim mDC As Long, mBitmap As Long
Private Sub Form_Click()
    Unload Me
End Sub
Private Sub Form_Load()

    Dim mRGN As Long, Cnt As Long, mBrush As Long, R As RECT
    'Create a device context, compatible with the screen
    mDC = CreateCompatibleDC(GetDC(0))
    'Create a bitmap, compatible with the screen
    mBitmap = CreateCompatibleBitmap(GetDC(0), Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY)
    'Select the bitmap nito the device context
    SelectObject mDC, mBitmap
    'Set the bitmap's backmode to transparent
    SetBkMode mDC, TRANSPARENT
    'Set the rectangles' values
    SetRect R, 0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY
    'Fill the rect with the default window-color
    FillRect mDC, R, GetSysColorBrush(COLOR_WINDOW)

    For Cnt = 0 To 350 Step 30
        'Select the new font into the form's device context and delete the old font
        DeleteObject SelectObject(mDC, CreateMyFont(24, Cnt))
        'Print some text
        TextOut mDC, (Me.Width / Screen.TwipsPerPixelX) / 2, (Me.Height / Screen.TwipsPerPixelY) / 2, Message, Len(Message)
    Next Cnt

    'Create an elliptical region
    mRGN = CreateEllipticRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY)
    'Set the window region
    SetWindowRgn Me.hWnd, mRGN, True

    'delete our elliptical region
    DeleteObject mRGN
End Sub
Function CreateMyFont(nSize As Integer, nDegrees As Long) As Long
    'Create a specified font
    CreateMyFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(GetDC(0), LOGPIXELSY), 72), 0, nDegrees * 10, 0, FW_NORMAL, False, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, "Times New Roman")
End Function
Private Sub Form_Paint()
    'Copy the picture to the form
    BitBlt Me.hdc, 0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, mDC, 0, 0, vbSrcCopy
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'clean up
    DeleteDC mDC
    DeleteObject mBitmap
End Sub

