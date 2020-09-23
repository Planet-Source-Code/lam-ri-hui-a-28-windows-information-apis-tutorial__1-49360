VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Colors Demo"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Const COLOR_SCROLLBAR = 0 'The Scrollbar colour
Const COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
Const COLOR_MENU = 4 'Menu
Const COLOR_WINDOW = 5 'Windows background
Const COLOR_WINDOWFRAME = 6 'Window frame
Const COLOR_MENUTEXT = 7 'Window Text
Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
Const COLOR_CAPTIONTEXT = 9 'Text in window caption
Const COLOR_ACTIVEBORDER = 10 'Border of active window
Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Const COLOR_HIGHLIGHT = 13 'Selected item background
Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
Const COLOR_BTNFACE = 15 'Button
Const COLOR_BTNSHADOW = 16 '3D shading of button
Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
Const COLOR_BTNTEXT = 18 'Button text
Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
Const COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color
Const COLOR_2NDINACTIVECAPTION = 28 'Win98 only: 2nd inactive window color
Private Sub Form_Load()

    'Get the caption's active color
    col& = GetSysColor(COLOR_ACTIVECAPTION)
    'Change the active caption's color to red
    t& = SetSysColors(1, COLOR_ACTIVECAPTION, RGB(255, 0, 0))
    MsgBox "The old title bar color was" + Str$(col&) + " and is now" + Str$(GetSysColor(COLOR_ACTIVECAPTION))
End Sub

