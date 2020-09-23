VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "User Name Extended Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum EXTENDED_NAME_FORMAT
    NameUnknown = 0
    NameFullyQualifiedDN = 1
    NameSamCompatible = 2
    NameDisplay = 3
    NameUniqueId = 6
    NameCanonical = 7
    NameUserPrincipal = 8
    NameCanonicalEx = 9
    NameServicePrincipal = 10
End Enum
Private Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long
Private Sub Form_Load()

    Dim sBuffer As String, Ret As Long
    sBuffer = String(256, 0)
    Ret = Len(sBuffer)
    If GetUserNameEx(NameSamCompatible, sBuffer, Ret) <> 0 Then
        MsgBox "Username: " + Left$(sBuffer, Ret)
    Else
        MsgBox "Error while retrieving the username"
    End If
End Sub

