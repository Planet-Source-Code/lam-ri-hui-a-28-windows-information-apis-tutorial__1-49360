VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Set Computer Name Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Sub Form_Load()
    Dim sNewName As String
    'Ask for a new computer name
    sNewName = InputBox("Please enter a new computer name.")
    'Set the new computer name
    SetComputerName sNewName
    MsgBox "Computername set to " + sNewName
End Sub

