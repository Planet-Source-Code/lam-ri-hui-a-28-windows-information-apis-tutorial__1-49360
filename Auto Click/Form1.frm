VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Auto Click Menu Demo"
   ClientHeight    =   3060
   ClientLeft      =   165
   ClientTop       =   930
   ClientWidth     =   4680
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
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menu1 
      Caption         =   "File"
      Begin VB.Menu menu2 
         Caption         =   "Click"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Const MOUSEEVENTF_ABSOLUTE = &H8000 ' absolute move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 ' left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 ' left button up
Private Const MOUSEEVENTF_MOVE = &H1 ' mouse move
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0 'X Size of screen
Const SM_CYSCREEN = 1 'Y Size of Screen

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim mWnd As Long
    mWnd = Me.hwnd
    
    Dim hMenu As Long, hSubMenu As Long

    hMenu = GetMenu(mWnd) 'Get the Menu of the Window(MenuBar)
    ClickMenuItem mWnd, hMenu, 0 'Click on the first SubMenu
    hSubMenu = GetSubMenu(hMenu, 0) 'Get its submenu
    ClickMenuItem mWnd, hSubMenu, 0 'Click on the first MenuItem of the Submenu
    
End Sub


Private Sub ScreenToAbsolute(lpPoint As POINTAPI)
lpPoint.x = lpPoint.x * (&HFFFF& / GetSystemMetrics(SM_CXSCREEN))
lpPoint.y = lpPoint.y * (&HFFFF& / GetSystemMetrics(SM_CYSCREEN))
End Sub

Private Sub Click(p As POINTAPI)
'p.X and p.Y in absolute coordinates
'Put the mouse on the point

mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, p.x, p.y, 0, GetMessageExtraInfo()
'Mouse Down
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, GetMessageExtraInfo()
'Mouse Up
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, GetMessageExtraInfo()
End Sub

Private Sub ClickMenuItem(ByVal mWnd As Long, ByVal hMenu As Long, ByVal Pos As Long)
Dim ret As Long
Dim r As RECT, p As POINTAPI
ret = GetMenuItemRect(mWnd, hMenu, Pos, r)
If ret = 0 Then Exit Sub
p.x = (r.Left + r.Right) / 2
p.y = (r.Top + r.Bottom) / 2
ScreenToAbsolute p
'Click on p
Click p
End Sub

Private Sub Form_Load()
Dim mWnd As Long, p As POINTAPI
mWnd = Me.hwnd
Dim hMenu As Long, hSubMenu As Long
hMenu = GetMenu(mWnd) 'Get the Menu of the Window(MenuBar)
ClickMenuItem mWnd, hMenu, 0 'Click on the first SubMenu
hSubMenu = GetSubMenu(hMenu, 0) 'Get its submenu
ClickMenuItem mWnd, hSubMenu, 0 'Click on the first MenuItem of the Submenu
p.x = &HFFFF& / 2
p.y = &HFFFF& / 2
Click p
Me.AutoRedraw = True
Me.BackColor = vbWhite
Print "Press any key"
End Sub

Private Sub menu2_Click()
MsgBox "Click"
End Sub

