VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Resume Demo"
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
Private Declare Function IsSystemResumeAutomatic Lib "kernel32" () As Long
Private Sub Form_Load()
    If IsSystemResumeAutomatic <> 0 Then
        MsgBox "The system was restored to the working state automatically and the user is not active."
    Else
        MsgBox "The system doesn't support automatic system restore."
    End If
End Sub

