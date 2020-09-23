VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Directory Demo"
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
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Sub Form_Load()
    'KPD-Team 1998

    Dim sSave As String, Ret As Long
    'Create a buffer
    sSave = Space(255)
    'Get the system directory
    Ret = GetSystemDirectory(sSave, 255)
    'Remove all unnecessary chr$(0)'s
    sSave = Left$(sSave, Ret)
    'Show the windows directory
    MsgBox "Windows System directory: " + sSave
End Sub

