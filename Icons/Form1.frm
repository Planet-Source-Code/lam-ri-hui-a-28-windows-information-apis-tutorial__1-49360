VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Icons Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project needs a PictureBox, called 'Picture1'

'In general section
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Sub Form_Load()


    Dim Path As String, strSave As String
    'Create a buffer string
    strSave = String(200, Chr$(0))
    'Get the windows directory and append '\REGEdit.exe' to it
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\REGEdit.exe"
    'No pictures
    Picture1.Picture = LoadPicture()
    'Set graphicmode to 'persistent
    Picture1.AutoRedraw = True
    'Extract the icon from REGEdit
    return1& = ExtractIcon(Me.hWnd, Path, 2)
    'Draw the icon on the form
    return2& = DrawIcon(Picture1.hdc, 0, 0, return1&)
End Sub

