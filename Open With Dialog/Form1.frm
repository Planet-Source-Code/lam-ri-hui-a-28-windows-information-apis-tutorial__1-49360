VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Show OpenWith Dialog Demo"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Const SE_ERR_NOASSOC = 31
Const sOperation As String = "open"     ' Constants for shell operations
Const sRun As String = "RUNDLL32.EXE"
Const sParameters As String = "shell32.dll,OpenAs_RunDLL "
Private Function shelldoc(sfile As String)
    Dim sPath As String, RetVal As Long, _
    lRet As Long
    lRet = ShellExecute(GetDesktopWindow(), sOperation, sfile, _
                        vbNullString, vbNullString, SW_SHOWNORMAL)
    If lRet = SE_ERR_NOASSOC Then ' No association exists
        'Create a buffer
        sPath = Space(255)
        'Get the system directory
        RetVal = GetSystemDirectory(sPath, 255)
        'Remove all unnecessary chr$(0)'s
        'and move on the stack
        sPath = Left$(sPath, RetVal)
    
        lRet = ShellExecute(GetDesktopWindow(), "open", sRun, _
                            sParameters + sfile, sPath, SW_SHOWNORMAL)
    End If
End Function
Private Sub Command1_Click()
    ' Change the file extensions so that one
    ' has a program associated with it and the
    ' other does not.
    Call shelldoc(App.Path & "\" & "test.txt")
    Call shelldoc(App.Path & "\" & "test.sarsaparilla")
End Sub

