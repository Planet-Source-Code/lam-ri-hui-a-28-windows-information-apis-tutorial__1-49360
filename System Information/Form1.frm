VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Information Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6225
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
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
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
Private Sub Form_Load()
    Dim SInfo As SYSTEM_INFO

    'Set the graphical mode to persistent
    Me.AutoRedraw = True
    'Get the system information
    GetSystemInfo SInfo
    'Print it to the form
    Me.Print "Number of procesor:" + Str$(SInfo.dwNumberOrfProcessors)
    Me.Print "Processor:" + Str$(SInfo.dwProcessorType)
    Me.Print "Low memory address:" + Str$(SInfo.lpMinimumApplicationAddress)
    Me.Print "High memory address:" + Str$(SInfo.lpMaximumApplicationAddress)
End Sub

