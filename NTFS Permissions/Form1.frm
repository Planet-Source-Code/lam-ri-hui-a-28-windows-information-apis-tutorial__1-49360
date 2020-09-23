VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "NTFS Permissions Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Example from MSDN (Q240176)
'The following code changes permissions on a folder to Add & Read or Change.
'The folder needs to be created on an NTFS partition.
'You need to be an Administrator on the machine in question and have read/write
'(READ_CONTROL and WRITE_DAC) access to the file or directory.

Private Sub Command1_Click()
    Dim sUserName As String
    Dim sFolderName As String
    sUserName = Trim$(CStr(Text2.Text))
    sFolderName = Trim$(CStr(Text1.Text))
    SetAccess sUserName, sFolderName, GENERIC_READ Or GENERIC_EXECUTE Or DELETE Or GENERIC_WRITE
End Sub
Private Sub Command2_Click()
    Dim sUserName As String
    Dim sFolderName As String
    sUserName = Trim$(Text2.Text)
    sFolderName = Trim$(Text1.Text)
    SetAccess sUserName, sFolderName, GENERIC_EXECUTE Or GENERIC_READ
End Sub
Private Sub Form_Load()
    Text1.Text = "enter folder name"
    Text2.Text = "enter username"
    Command1.Caption = "Change"
    Command2.Caption = "Read && Add"
End Sub



