VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Using Small Fonts Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MM_TEXT = 1
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Function SmallFonts() As Boolean
   Dim hdc As Long, hwnd As Long
   Dim PrevMapMode As Long, tm As TEXTMETRIC

   ' Set the default return value to small fonts
   SmallFonts = True

   ' Get the handle of the desktop window
   hwnd = GetDesktopWindow()

   ' Get the device context for the desktop
   hdc = GetWindowDC(hwnd)
   If hdc Then
      ' Set the mapping mode to pixels
      PrevMapMode = SetMapMode(hdc, MM_TEXT)

      ' Get the size of the system font
      GetTextMetrics hdc, tm

      ' Set the mapping mode back to what it was
      PrevMapMode = SetMapMode(hdc, PrevMapMode)

      ' Release the device context
      ReleaseDC hwnd, hdc

      ' If the system font is more than 16 pixels high,
      ' then large fonts are being used
      If tm.tmHeight > 16 Then SmallFonts = False
   End If
End Function
Private Sub Form_Load()

    MsgBox "Using small fonts: " + Str$(SmallFonts)
End Sub

