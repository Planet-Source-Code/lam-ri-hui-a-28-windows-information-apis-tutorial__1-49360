VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Window(s) Information Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Sub Form_Load()

    Timer1.Interval = 100
    Timer1.Enabled = True
    Dim strTemp As String, strUserName As String
    'Create a buffer
    strTemp = String(100, Chr$(0))
    'Get the temporary path
    GetTempPath 100, strTemp
    'strip the rest of the buffer
    strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)

    'Create a buffer
    strUserName = String(100, Chr$(0))
    'Get the username
    GetUserName strUserName, 100
    'strip the rest of the buffer
    strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)

    'Show the temppath and the username
    MsgBox "Hello " + strUserName + Chr$(13) + "The temp. path is " + strTemp
End Sub
Private Sub Timer1_Timer()
    Dim Boo As Boolean
    'Check if this form is minimized
    Boo = IsIconic(Me.hwnd)
    'Update the form's caption
    Me.Caption = Me.Caption & " - Form minimized: " + Str$(Boo)
End Sub

