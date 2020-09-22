VERSION 5.00
Begin VB.Form mouse1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Lock"
   ClientHeight    =   2055
   ClientLeft      =   2550
   ClientTop       =   2835
   ClientWidth     =   7770
   Icon            =   "Mouse1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7770
   Begin VB.CommandButton Command3 
      Caption         =   "Go To Password Generator"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4920
      Top             =   -120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Please Enter Then Verify A Password"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   7575
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Please Enter Your Name"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "mouse1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text2.Text <> Text3.Text Or Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("Error:  Be sure you have entered and varified your password correcly and have not left anything blank")
Else
mouse2.Show
mouse1.Hide
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
passgen.Show
End Sub

Private Sub Command4_Click()
Form1.Show
End Sub

Private Sub Form_Load()

End Sub
