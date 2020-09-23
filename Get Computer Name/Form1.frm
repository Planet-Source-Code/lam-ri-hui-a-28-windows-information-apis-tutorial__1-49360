VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get Computer Name Demo"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub Form_Load()
    Dim dwLen As Long
    Dim strString As String
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    MsgBox strString, , "Computer Name"
End Sub
