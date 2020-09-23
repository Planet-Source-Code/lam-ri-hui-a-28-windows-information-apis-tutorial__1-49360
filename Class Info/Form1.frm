VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Class Info Demo"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Private Sub Form_Paint()

    Dim WC As WNDCLASS
    'Get class info
    GetClassInfo ByVal 0&, "BUTTON", WC
    'Clear the form
    Me.Cls
    'Print the retrieved information to the form
    Me.Print "The button's default background is set to color-number:" + Str$(GetSysColor(WC.hbrBackground))
End Sub

