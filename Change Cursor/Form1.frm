VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Change Cursor Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Change Cursor to Banana"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore Default"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Const OCR_NORMAL As Long = 32512
Private currenthcurs As Long
Private tempcurs As Long
Private newhcurs As Long
Private Sub Command1_Click()
    Dim myDir As String
    Dim lDir As Long
    myDir = Space(255)
    currenthcurs = GetCursor()
    tempcurs = CopyIcon(currenthcurs)
    lDir = GetWindowsDirectory(myDir, 255)
    myDir = Left$(myDir, lDir) & "\cursors\banana.ani"
    newhcurs = LoadCursorFromFile(myDir)
    Call SetSystemCursor(newhcurs, OCR_NORMAL)
End Sub
Private Sub Command2_Click()
    Call SetSystemCursor(tempcurs, OCR_NORMAL)
End Sub

