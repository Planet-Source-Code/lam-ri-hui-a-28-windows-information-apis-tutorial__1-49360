VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Shutdown Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here to Shutdown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Shutdown Flags
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const SE_PRIVILEGE_ENABLED = &H2
Const TokenPrivileges = 3
Const TOKEN_ASSIGN_PRIMARY = &H1
Const TOKEN_DUPLICATE = &H2
Const TOKEN_IMPERSONATE = &H4
Const TOKEN_QUERY = &H8
Const TOKEN_QUERY_SOURCE = &H10
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_ADJUST_GROUPS = &H40
Const TOKEN_ADJUST_DEFAULT = &H80
Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Const ANYSIZE_ARRAY = 1
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Type Luid
    lowpart As Long
    highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    'pLuid As Luid
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Public Function InitiateShutdownMachine(ByVal Machine As String, Optional Force As Variant, Optional Restart As Variant, Optional AllowLocalShutdown As Variant, Optional Delay As Variant, Optional Message As Variant) As Boolean
    Dim hProc As Long
    Dim OldTokenStuff As TOKEN_PRIVILEGES
    Dim OldTokenStuffLen As Long
    Dim NewTokenStuff As TOKEN_PRIVILEGES
    Dim NewTokenStuffLen As Long
    Dim pSize As Long
    If IsMissing(Force) Then Force = False
    If IsMissing(Restart) Then Restart = True
    If IsMissing(AllowLocalShutdown) Then AllowLocalShutdown = False
    If IsMissing(Delay) Then Delay = 0
    If IsMissing(Message) Then Message = ""
    'Make sure the Machine-name doesn't start with '\\'
    If InStr(Machine, "\\") = 1 Then
        Machine = Right(Machine, Len(Machine) - 2)
    End If
    'check if it's the local machine that's going to be shutdown
    If (LCase(GetMyMachineName) = LCase(Machine)) Then
        'may we shut this computer down?
        If AllowLocalShutdown = False Then Exit Function
        'open access token
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hProc) = 0 Then
            MsgBox "OpenProcessToken Error: " & GetLastError()
            Exit Function
        End If
        'retrieve the locally unique identifier to represent the Shutdown-privilege name
        If LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, OldTokenStuff.Privileges(0).pLuid) = 0 Then
            MsgBox "LookupPrivilegeValue Error: " & GetLastError()
            Exit Function
        End If
        NewTokenStuff = OldTokenStuff
        NewTokenStuff.PrivilegeCount = 1
        NewTokenStuff.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        NewTokenStuffLen = Len(NewTokenStuff)
        pSize = Len(NewTokenStuff)
        'Enable shutdown-privilege
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen) = 0 Then
            MsgBox "AdjustTokenPrivileges Error: " & GetLastError()
            Exit Function
        End If
        'initiate the system shutdown
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
        NewTokenStuff.Privileges(0).Attributes = 0
        'Disable shutdown-privilege
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, Len(NewTokenStuff), OldTokenStuff, Len(OldTokenStuff)) = 0 Then
            Exit Function
        End If
    Else
        'initiate the system shutdown
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
    End If
    InitiateShutdownMachine = True
End Function
Function GetMyMachineName() As String
    Dim sLen As Long
    'create a buffer
    GetMyMachineName = Space(100)
    sLen = 100
    'retrieve the computer name
    If GetComputerName(GetMyMachineName, sLen) Then
        GetMyMachineName = Left(GetMyMachineName, sLen)
    End If
End Function

Private Sub Command1_Click()
 InitiateShutdownMachine GetMyMachineName, True, True, True, 60, "You initiated a system shutdown..."
End Sub

Private Sub Form_Load()

   
End Sub

