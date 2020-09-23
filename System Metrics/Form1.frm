VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Metrics Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In general section
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0 'X Size of screen
Const SM_CYSCREEN = 1 'Y Size of Screen
Const SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
Const SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
Const SM_CYCAPTION = 4 'Height of windows caption
Const SM_CXBORDER = 5 'Width of no-sizable borders
Const SM_CYBORDER = 6 'Height of non-sizable borders
Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Const SM_CYDLGFRAME = 8 'Height of dialog box borders
Const SM_CYVTHUMB = 9 'Height of scroll box on horizontal scroll bar
Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Const SM_CXICON = 11 'Width of standard icon
Const SM_CYICON = 12 'Height of standard icon
Const SM_CXCURSOR = 13 'Width of standard cursor
Const SM_CYCURSOR = 14 'Height of standard cursor
Const SM_CYMENU = 15 'Height of menu
Const SM_CXFULLSCREEN = 16 'Width of client area of maximized window
Const SM_CYFULLSCREEN = 17 'Height of client area of maximized window
Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Const SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
Const SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
Const SM_DEBUG = 22 'True if deugging version of windows is running
Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Const SM_CXMIN = 28 'Minimum width of window
Const SM_CYMIN = 29 'Minimum height of window
Const SM_CXSIZE = 30 'Width of title bar bitmaps
Const SM_CYSIZE = 31 'height of title bar bitmaps
Const SM_CXMINTRACK = 34 'Minimum tracking width of window
Const SM_CYMINTRACK = 35 'Minimum tracking height of window
Const SM_CXDOUBLECLK = 36 'double click width
Const SM_CYDOUBLECLK = 37 'double click height
Const SM_CXICONSPACING = 38 'width between desktop icons
Const SM_CYICONSPACING = 39 'height between desktop icons
Const SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
Const SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
Const SM_DBCSENABLED = 42 'True if double byte characteds are enabled
Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Const SM_CMETRICS = 44 'Number of system metrics
Const SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
Const SM_CXMAXIMIZED = 61 'default width of win95 maximised window
Const SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
Const SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
Const SM_CXMENUSIZE = 54 'width of button on menu bar
Const SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
Const SM_CYMAXIMIZED = 62 'default height of win95 maximised window
Const SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
Const SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
Const SM_CYMENUSIZE = 55 'height of button on menu bar
Const SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
Const SM_CYSMCAPTION = 51 'height of windows 95 small caption
Const SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
Const SM_NETWORK = 63 'bit o is set if a network is present. Const SM_SECURE = 44 'True if security is present on windows 95 system
Const SM_SLOWMACHINE = 73 'true if machine is too slow to run win95.
Private Sub Form_Load()

    'Set the graphic mode to persistent
    Me.AutoRedraw = True
    'retrieve information and print it to the form
    Me.Print "Number of mouse buttons:" + Str$(GetSystemMetrics(SM_CMOUSEBUTTONS))
    Me.Print "Screen X:" + Str$(GetSystemMetrics(SM_CXSCREEN))
    Me.Print "Screen Y:" + Str$(GetSystemMetrics(SM_CYSCREEN))
    Me.Print "Height of windows caption:" + Str$(GetSystemMetrics(SM_CYCAPTION))
    Me.Print "Width between desktop icons:" + Str$(GetSystemMetrics(SM_CXICONSPACING))
    Me.Print "Maximum width when resizing a window:" + Str$(GetSystemMetrics(SM_CYMAXTRACK))
    Me.Print "Is machine is too slow to run windows?" + Str$(GetSystemMetrics(SM_SLOWMACHINE))
End Sub

