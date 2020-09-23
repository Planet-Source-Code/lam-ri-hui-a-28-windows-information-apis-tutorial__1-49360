VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WOW64 Process Demo"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub GetNativeSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Sub Form_Load()

    Dim Ret As Long
    IsWow64Process GetCurrentProcess, Ret
    If Ret = 0 Then
        MsgBox "This application is not running on an x86 emulator for a 64-bit computer!"
    Else
        Dim SysInfo64 As SYSTEM_INFO
        GetNativeSystemInfo SysInfo64
        MsgBox "Number of processors on your 64-bit system: " + CStr(SysInfo64.dwNumberOrfProcessors)
    End If
End Sub

