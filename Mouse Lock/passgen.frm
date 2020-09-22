VERSION 5.00
Begin VB.Form passgen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Password Generater"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "passgen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   2640
   End
   Begin VB.CommandButton Command4 
      Caption         =   "STOP!!"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Super Password Generator"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Password"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generated Password"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   6735
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6255
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Choose Number Of Digits The Password Will Have"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "passgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function GeneratePassword(Length As Integer) As String

Dim blnOnVowel As Boolean
Dim strTempLetter As String
Dim strPassword As String

For i = 1 To Length
If blnOnVowel = False Then
   strTempLetter = Choose(GetRandomNumber(1, 17), _
   "B", "D", "F", "G", "H", "J", "K", "L", "M", _
   "N", "P", "R", "S", "T", "V", "W", "Y")
   strPassword = strPassword & strTempLetter
  blnOnVowel = True
Else
  strTempLetter = Choose(GetRandomNumber(1, 5), _
  "A", "E", "I", "O", "U")
  
  strPassword = strPassword & strTempLetter
  blnOnVowel = False
End If
Next i
GeneratePassword = strPassword
End Function

Public Function GetRandomNumber(Upper As Integer, _
Lower As Integer) As Integer
Randomize
GetRandomNumber = Int((Upper - Lower + 1) * Rnd + Lower)
End Function




Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox ("Please select a number of digits for the password")
Else
Text1.Text = GeneratePassword(Combo1.Text)
End If
End Sub

Private Sub Command2_Click()
passgen.Hide
Unload passgen
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Command4.Visible = True
Command3.Visible = False
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Command4.Visible = False
Command3.Visible = True
End Sub

Private Sub Form_Load()
Text1.FontBold = True
Text1.FontSize = 16
Combo1.AddItem 4
Combo1.AddItem 5
Combo1.AddItem 6
Combo1.AddItem 7
Combo1.AddItem 8
Combo1.AddItem 9
Combo1.AddItem 10
Combo1.AddItem 11
Combo1.AddItem 12
Combo1.AddItem 13
Combo1.AddItem 14
Combo1.AddItem 15
End Sub

Private Sub Timer1_Timer()
Text1.Text = GeneratePassword(Rnd * 15)
End Sub

Private Sub Timer2_Timer()
Text1.Text = GeneratePassword(Rnd * 15)
End Sub

Private Sub Timer3_Timer()
Text1.Text = GeneratePassword(Rnd * 15)
End Sub
