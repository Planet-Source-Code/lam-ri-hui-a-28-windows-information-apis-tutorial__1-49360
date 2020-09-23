VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "User Directories Demo"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const TOKEN_QUERY = (&H8)
Private Declare Function GetAllUsersProfileDirectory Lib "userenv.dll" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetDefaultUserProfileDirectory Lib "userenv.dll" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetProfilesDirectory Lib "userenv.dll" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Sub Form_Load()

    Dim sBuffer As String, Ret As Long, hToken As Long
    'set the graphics mode of this form to 'persistent'
    Me.AutoRedraw = True
    'create a string buffer
    sBuffer = String(255, 0)
    'retrieve the all users profile directory
    GetAllUsersProfileDirectory sBuffer, 255
    'show the result
    Me.Print StripTerminator(sBuffer)
    'create a string buffer
    sBuffer = String(255, 0)
    'retrieve the user profile directory
    GetDefaultUserProfileDirectory sBuffer, 255
    'show the result
    Me.Print StripTerminator(sBuffer)
    'create a string buffer
    sBuffer = String(255, 0)
    'retrieve the profiles directory
    GetProfilesDirectory sBuffer, 255
    'show the result
    Me.Print StripTerminator(sBuffer)
    'create a string buffer
    sBuffer = String(255, 0)
    'open the token of the current process
    OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
    'retrieve this users profile directory
    GetUserProfileDirectory hToken, sBuffer, 255
    'show the result
    Me.Print StripTerminator(sBuffer)
End Sub
'strips off the trailing Chr$(0)'s
Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

