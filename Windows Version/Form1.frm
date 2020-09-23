VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Windows Version Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Sub Form_Load()
    Dim OSInfo As OSVERSIONINFO, PId As String

    'Set the graphical mode to persistent
    Me.AutoRedraw = True
    'Set the structure size
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'Get the Windows version
    Ret& = GetVersionEx(OSInfo)
    'Chack for errors
    If Ret& = 0 Then MsgBox "Error Getting Version Information": Exit Sub
    'Print the information to the form
    Select Case OSInfo.dwPlatformId
        Case 0
            PId = "Windows 32s "
        Case 1
            PId = "Windows 95/98"
        Case 2
            PId = "Windows NT "
    End Select
    Print "OS: " + PId
    Print "Win version:" + Str$(OSInfo.dwMajorVersion) + "." + LTrim(Str(OSInfo.dwMinorVersion))
    Print "Build: " + Str(OSInfo.dwBuildNumber)
End Sub

