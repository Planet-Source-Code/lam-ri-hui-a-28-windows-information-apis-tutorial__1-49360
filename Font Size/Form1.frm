VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Font Size Demo"
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
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Const LOGPIXELSX = 88
Public Function IsScreenFontSmall() As Boolean
    Dim hWndDesk As Long, hDCDesk As Long
    Dim logPix As Long, r As Long
    'Get the handle of the desktop window
    hWndDesk = GetDesktopWindow()
    'Get the desktop window's device context
    hDCDesk = GetDC(hWndDesk)
    'Get the width of the screen
    logPix = GetDeviceCaps(hDCDesk, LOGPIXELSX)
    'Release the device context
    r = ReleaseDC(hWndDesk, hDCDesk)
    IsScreenFontSmall = (logPix = 96)
End Function
Private Sub Form_Load()

    If IsScreenFontSmall = True Then
        MsgBox "You're using Small Fonts"
    Else
        MsgBox "You're using Large Fonts"
    End If
End Sub

