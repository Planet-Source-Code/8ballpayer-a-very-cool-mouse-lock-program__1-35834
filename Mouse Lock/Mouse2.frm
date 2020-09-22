VERSION 5.00
Begin VB.Form mouse2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2685
   ClientLeft      =   2505
   ClientTop       =   2505
   ClientWidth     =   6270
   Icon            =   "Mouse2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   5
      Left            =   600
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   5520
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   120
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      Caption         =   "Number Of Incorrect Password Attempts"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Password"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "mouse2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Agent1_ActivateInput(ByVal CharacterID As String)

End Sub

Private Sub Command1_Click()
Load mouse1
If Text1.Text = mouse1.Text2.Text Then
SetOnTop mouse2.hwnd, False
MsgBox ("Welcome " + mouse1.Text1.Text)
Timer3.Enabled = False
DisableTrap mouse2
End
Else
Label1.Caption = "Incorrect Password"
Timer1.Enabled = True
Label2.Caption = Label2.Caption + 1
EnableTrap mouse2
End If
End Sub
Public Sub HideTask(Hide As Boolean)
Dim lHandle As Long
Dim lService As Long
lHandle = GetCurrentProcessId()
lService = RegisterServiceProcess(lHandle, Abs(Hide))
End Sub
Private Sub Form_Load()
SetOnTop mouse2.hwnd, True
 HideTask True
EnableTrap mouse2
Load mouse1
Label3.Caption = "Please enter a password as proof that you are " + mouse1.Text1.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
EnableTrap mouse2
End Sub


Private Sub Timer1_Timer()
Label1.Caption = "Enter Password"

End Sub


Private Sub Timer2_Timer()
If GetAsyncKeyState(VK_F9) Then
    Label1.Caption = "Blah"
End If
End Sub

Private Sub Timer3_Timer()
EnableTrap mouse2
End Sub
