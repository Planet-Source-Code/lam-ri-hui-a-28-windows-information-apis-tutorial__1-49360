VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sys Colors Demo"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const COLOR_SCROLLBAR = 0
Const COLOR_BACKGROUND = 1
Const COLOR_ACTIVECAPTION = 2
Const COLOR_INACTIVECAPTION = 3
Const COLOR_MENU = 4
Const COLOR_WINDOW = 5
Const COLOR_WINDOWFRAME = 6
Const COLOR_MENUTEXT = 7
Const COLOR_WINDOWTEXT = 8
Const COLOR_CAPTIONTEXT = 9
Const COLOR_ACTIVEBORDER = 10
Const COLOR_INACTIVEBORDER = 11
Const COLOR_APPWORKSPACE = 12
Const COLOR_HIGHLIGHT = 13
Const COLOR_HIGHLIGHTTEXT = 14
Const COLOR_BTNFACE = 15
Const COLOR_BTNSHADOW = 16
Const COLOR_GRAYTEXT = 17
Const COLOR_BTNTEXT = 18
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Dim SavedColors(18) As Long, IndexArray(18) As Long, NewColors(18) As Long
Private Sub Form_Load()

    ' Save current system colors:
    For i = 0 To 18
        SavedColors(i) = GetSysColor(i)
    Next i

    ' Change all display elements:
    For i = 0 To 18
        Randomize Timer
        NewColors(i) = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        IndexArray(i) = i
    Next i
    SetSysColors 19, IndexArray(0), NewColors(0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' Restore system colors:
    SetSysColors 19, IndexArray(0), SavedColors(0)
End Sub

