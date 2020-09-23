VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "WallPaper Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CDBox 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1

Private Sub Command1_Click()

    'Set the commondialogbox' title
    CDBox.DialogTitle = "Choose a bitmap"
    'Set the filter
    CDBox.Filter = "Windows Bitmaps (*.BMP)|*.bmp|All Files (*.*)|*.*"
    'Show the 'Open File'-dialog
    CDBox.ShowOpen
    'Change the desktop's background
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0, CDBox.FileName, SPIF_UPDATEINIFILE
End Sub

Private Sub Form_Load()
    Command1.Caption = "Set Wallpaper"
End Sub

