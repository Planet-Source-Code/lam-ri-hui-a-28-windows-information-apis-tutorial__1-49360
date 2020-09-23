VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code retrieves information about the fonts
'used in the window menus and captions
Option Explicit
Const SPI_GETNONCLIENTMETRICS = 41
Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfEscapement As Long
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfWidth As Long
    lfWeight As Long
    lfItalic As Byte
    lfCharSet As Byte
    lfClipPrecision As Byte
    lfOutPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfOrientation As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Sub LastError()
    Dim Buffer As String
    Buffer = Space(200)
    SetLastError ERROR_BAD_USERNAME
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, Buffer, 200, ByVal 0&
    MsgBox Buffer
End Sub
Private Sub Form_Load()
    Dim ncm As NONCLIENTMETRICS, res&, strPuffer$, i%
    ncm.cbSize = 340
    res = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, ncm.cbSize, ncm, 0)
    If res = 0 Then LastError: Exit Sub
    Debug.Print "MenuFont.Height    " & CInt(-0.75 * ncm.lfMenuFont.lfHeight)
    Debug.Print "MenuFont.Weight    " & ncm.lfMenuFont.lfWeight     '400 = Normal, 700 = Bold
    Debug.Print "MenuFont.Italic    " & ncm.lfMenuFont.lfItalic             '0 = False, 1 = True
    strPuffer = StrConv(ncm.lfMenuFont.lfFaceName(), vbUnicode)
    Debug.Print "MenuFont    " & strPuffer
    Debug.Print "CaptionFont.Height    " & CInt(-0.75 * ncm.lfCaptionFont.lfHeight)
    Debug.Print "CaptionFont.Weight    " & ncm.lfCaptionFont.lfWeight
    Debug.Print "CaptionFont.Italic    " & ncm.lfCaptionFont.lfItalic
    strPuffer = StrConv(ncm.lfCaptionFont.lfFaceName(), vbUnicode)
    Debug.Print "CaptionFont    " & strPuffer

End Sub
