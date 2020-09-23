VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get Windows Version Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVersion Lib "kernel32" () As Long
Public Function GetWinVersion() As String
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function
Private Sub Form_Load()

    MsgBox "Windows version: " + GetWinVersion
End Sub

