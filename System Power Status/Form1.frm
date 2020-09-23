VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Power Status Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5790
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
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type
Private Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Sub Form_Paint()

    Dim SPS As SYSTEM_POWER_STATUS
    'get the battery powerstatus
    GetSystemPowerStatus SPS
    Me.AutoRedraw = True
    'show some information
    Select Case SPS.ACLineStatus
        Case 0
            Me.Print "AC power status: Offline"
        Case 1
            Me.Print "AC power status: OnLine"
        Case 2
            Me.Print "AC power status: Unknown"
    End Select
    Select Case SPS.BatteryFlag
        Case 1
            Me.Print "Battery charge status: High"
        Case 2
            Me.Print "Battery charge status: Low"
        Case 4
            Me.Print "Battery charge status: Critical"
        Case 8
            Me.Print "Battery charge status: Charging"
        Case 128
            Me.Print "Battery charge status: No system battery"
        Case 255
            Me.Print "Battery charge status: Unknown Status"
    End Select
End Sub

