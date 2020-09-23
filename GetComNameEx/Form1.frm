VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GetComputerNameEx Demo"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "The information are as shown in the immediate window"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum COMPUTER_NAME_FORMAT
    ComputerNameNetBIOS
    ComputerNameDnsHostname
    ComputerNameDnsDomain
    ComputerNameDnsFullyQualified
    ComputerNamePhysicalNetBIOS
    ComputerNamePhysicalDnsHostname
    ComputerNamePhysicalDnsDomain
    ComputerNamePhysicalDnsFullyQualified
    ComputerNameMax
End Enum
Private Declare Function GetComputerNameEx Lib "kernel32.dll" Alias "GetComputerNameExA" (ByVal NameType As COMPUTER_NAME_FORMAT, ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Sub Form_Load()

    ShowName ComputerNameNetBIOS, "NetBIOS name"
    ShowName ComputerNameDnsHostname, "DNS host name"
    ShowName ComputerNameDnsDomain, "DNS Domain"
    ShowName ComputerNameDnsFullyQualified, "Fully qualified DNS name"
    ShowName ComputerNamePhysicalNetBIOS, "Physical NetBIOS name"
End Sub
Private Sub ShowName(lIndex As COMPUTER_NAME_FORMAT, Description As String)
    Dim Ret As Long, sBuffer As String
    'create a buffer
    sBuffer = Space(256)
    Ret = Len(sBuffer)
    'retrieve the computer name
    If GetComputerNameEx(lIndex, sBuffer, Ret) <> 0 And Ret > 0 Then
        'show it
        Debug.Print Description + ": " + Left$(sBuffer, Ret)
    End If
End Sub

