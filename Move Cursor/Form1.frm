VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Move Cursor Demo"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   1920
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Dim P As POINTAPI
Private Sub Form_Load()


    Command1.Caption = "Screen Middle"
    Command2.Caption = "Form Middle"
    'API uses pixels
    Me.ScaleMode = vbPixels
End Sub
Private Sub Command1_Click()
    'Get information about the screen's width
    P.x = GetDeviceCaps(Form1.hdc, 8) / 2
    'Get information about the screen's height
    P.y = GetDeviceCaps(Form1.hdc, 10) / 2
    'Set the mouse cursor to the middle of the screen
    ret& = SetCursorPos(P.x, P.y)
End Sub
Private Sub Command2_Click()
    P.x = 0
    P.y = 0
    'Get information about the form's left and top
    ret& = ClientToScreen&(Form1.hwnd, P)
    P.x = P.x + Me.ScaleWidth / 2
    P.y = P.y + Me.ScaleHeight / 2
    'Set the cursor to the middle of the form
    ret& = SetCursorPos&(P.x, P.y)
End Sub

