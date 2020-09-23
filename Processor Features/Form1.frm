VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Processor Features Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const PF_FLOATING_POINT_PRECISION_ERRATA = 0
Private Const PF_FLOATING_POINT_EMULATED = 1
Private Const PF_COMPARE_EXCHANGE_DOUBLE = 2
Private Const PF_MMX_INSTRUCTIONS_AVAILABLE = 3
Private Const PF_XMMI_INSTRUCTIONS_AVAILABLE = 6
Private Const PF_3DNOW_INSTRUCTIONS_AVAILABLE = 7
Private Const PF_RDTSC_INSTRUCTION_AVAILABLE = 8
Private Const PF_PAE_ENABLED = 9
Private Declare Function IsProcessorFeaturePresent Lib "kernel32.dll" (ByVal ProcessorFeature As Long) As Long
Private Sub Form_Load()

    ShowFeature PF_FLOATING_POINT_PRECISION_ERRATA, "Floating point error"
    ShowFeature PF_FLOATING_POINT_EMULATED, "Floating-point operations emulated"
    ShowFeature PF_COMPARE_EXCHANGE_DOUBLE, "Compare and exchange double operation available"
    ShowFeature PF_MMX_INSTRUCTIONS_AVAILABLE, "MMX instructions available"
    ShowFeature PF_XMMI_INSTRUCTIONS_AVAILABLE, "XMMI instructions available"
    ShowFeature PF_3DNOW_INSTRUCTIONS_AVAILABLE, "3D-Now instructions available"
    ShowFeature PF_RDTSC_INSTRUCTION_AVAILABLE, "RDTSC instructions available"
    ShowFeature PF_PAE_ENABLED, "Processor is PAE-enabled"
End Sub
Private Sub ShowFeature(lIndex As Long, Description As String)
    If IsProcessorFeaturePresent(lIndex) = 0 Then
        Debug.Print Description + ": false"
    Else
        Debug.Print Description + ": true"
    End If
End Sub

