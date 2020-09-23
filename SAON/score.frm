VERSION 5.00
Begin VB.Form highscores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HighScore List"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   4680
      TabIndex        =   35
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play Game"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox newnick 
      Height          =   285
      Left            =   1320
      TabIndex        =   34
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox newscore 
      Height          =   285
      Left            =   2640
      TabIndex        =   33
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label s10 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   32
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label s9 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   31
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label s8 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   30
      Top             =   2280
      Width           =   90
   End
   Begin VB.Label s7 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   29
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label s6 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   28
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label s5 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   27
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label s4 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   26
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label s3 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   25
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label s2 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   24
      Top             =   840
      Width           =   90
   End
   Begin VB.Label s1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   23
      Top             =   600
      Width           =   90
   End
   Begin VB.Label c10 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   2760
      Width           =   180
   End
   Begin VB.Label c9 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   2520
      Width           =   180
   End
   Begin VB.Label c8 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   20
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label c7 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   19
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label c6 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   18
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label c5 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   17
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label c4 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   16
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label c3 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   15
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label c2 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   840
      Width           =   180
   End
   Begin VB.Label c1 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   720
      TabIndex        =   13
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      Caption         =   "10."
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "9."
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "8."
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "7."
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "6."
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "5."
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "4."
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "3."
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "2."
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "1."
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   240
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "High Scores:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "highscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
splash.Show
End Sub


Sub scoreset()

If Val(newscore.Text) > Val(s1.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = c7.Caption
s8.Caption = s7.Caption
c7.Caption = c6.Caption
s7.Caption = s6.Caption
c6.Caption = c5.Caption
s6.Caption = s5.Caption
c5.Caption = c4.Caption
s5.Caption = s4.Caption
c4.Caption = c3.Caption
s4.Caption = s3.Caption
c3.Caption = c2.Caption
s3.Caption = s2.Caption
c2.Caption = c1.Caption
s2.Caption = s1.Caption
c1.Caption = newnick.Text
s1.Caption = newscore.Text


ElseIf Val(newscore.Text) > Val(s2.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = c7.Caption
s8.Caption = s7.Caption
c7.Caption = c6.Caption
s7.Caption = s6.Caption
c6.Caption = c5.Caption
s6.Caption = s5.Caption
c5.Caption = c4.Caption
s5.Caption = s4.Caption
c4.Caption = c3.Caption
s4.Caption = s3.Caption
c3.Caption = c2.Caption
s3.Caption = s2.Caption
c2.Caption = newnick.Text
s2.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s3.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = c7.Caption
s8.Caption = s7.Caption
c7.Caption = c6.Caption
s7.Caption = s6.Caption
c6.Caption = c5.Caption
s6.Caption = s5.Caption
c5.Caption = c4.Caption
s5.Caption = s4.Caption
c4.Caption = c3.Caption
s4.Caption = s3.Caption
c3.Caption = newnick.Text
s3.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s4.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = c7.Caption
s8.Caption = s7.Caption
c7.Caption = c6.Caption
s7.Caption = s6.Caption
c6.Caption = c5.Caption
s6.Caption = s5.Caption
c5.Caption = c4.Caption
s5.Caption = s4.Caption
c4.Caption = newnick.Text
s4.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s5.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = c7.Caption
s8.Caption = s7.Caption
c7.Caption = c6.Caption
s7.Caption = s6.Caption
c6.Caption = c5.Caption
s6.Caption = s5.Caption
c5.Caption = newnick.Text
s5.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s6.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = c7.Caption
s8.Caption = s7.Caption
c7.Caption = c6.Caption
s7.Caption = s6.Caption
c6.Caption = newnick.Text
s6.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s7.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = c7.Caption
s8.Caption = s7.Caption
c7.Caption = newnick.Text
s7.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s8.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c8.Caption = newnick.Text
s8.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s9.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c9.Caption = c8.Caption
s9.Caption = s8.Caption
c9.Caption = newnick.Text
s9.Caption = newscore.Text

ElseIf Val(newscore.Text) > Val(s10.Caption) Then
c10.Caption = c9.Caption
s10.Caption = s9.Caption
c10.Caption = newnick.Text
s10.Caption = newscore.Text
End If

Exit Sub
End Sub


Private Sub Command3_Click()
DeleteSetting "SAON", "highscores", "nick1"
DeleteSetting "SAON", "highscores", "nick2"
DeleteSetting "SAON", "highscores", "nick3"
DeleteSetting "SAON", "highscores", "nick4"
DeleteSetting "SAON", "highscores", "nick5"
DeleteSetting "SAON", "highscores", "nick6"
DeleteSetting "SAON", "highscores", "nick7"
DeleteSetting "SAON", "highscores", "nick8"
DeleteSetting "SAON", "highscores", "nick9"
DeleteSetting "SAON", "highscores", "nick10"

DeleteSetting "SAON", "highscores", "score1"
DeleteSetting "SAON", "highscores", "score2"
DeleteSetting "SAON", "highscores", "score3"
DeleteSetting "SAON", "highscores", "score4"
DeleteSetting "SAON", "highscores", "score5"
DeleteSetting "SAON", "highscores", "score6"
DeleteSetting "SAON", "highscores", "score7"
DeleteSetting "SAON", "highscores", "score8"
DeleteSetting "SAON", "highscores", "score9"
DeleteSetting "SAON", "highscores", "score10"

s1.Caption = "0"
s2.Caption = "0"
s3.Caption = "0"
s4.Caption = "0"
s5.Caption = "0"
s6.Caption = "0"
s7.Caption = "0"
s8.Caption = "0"
s9.Caption = "0"
s10.Caption = "0"

c1.Caption = "0"
c2.Caption = "0"
c3.Caption = "0"
c4.Caption = "0"
c5.Caption = "0"
c6.Caption = "0"
c7.Caption = "0"
c8.Caption = "0"
c9.Caption = "0"
End Sub

Private Sub Form_Load()
If main.cheatfixer.Text = "" And main.debtcheck.Text = "paid" Then
newnick.Text = main.nickname.Caption
newscore.Text = main.totcash.Caption
loader
scoreset
saver
Unload main
Else
loader
Me.Show
MsgBox "Invalid Score, you Cheated. or Didnt pay off your debt.", vbExclamation, "Score Unlogged."
End If
End Sub

Sub saver()
SaveSetting "SAON", "highscores", "nick1", c1.Caption
SaveSetting "SAON", "highscores", "nick2", c2.Caption
SaveSetting "SAON", "highscores", "nick3", c3.Caption
SaveSetting "SAON", "highscores", "nick4", c4.Caption
SaveSetting "SAON", "highscores", "nick5", c5.Caption
SaveSetting "SAON", "highscores", "nick6", c6.Caption
SaveSetting "SAON", "highscores", "nick7", c7.Caption
SaveSetting "SAON", "highscores", "nick8", c8.Caption
SaveSetting "SAON", "highscores", "nick9", c9.Caption
SaveSetting "SAON", "highscores", "nick10", c10.Caption

SaveSetting "SAON", "highscores", "score1", s1.Caption
SaveSetting "SAON", "highscores", "score2", s2.Caption
SaveSetting "SAON", "highscores", "score3", s3.Caption
SaveSetting "SAON", "highscores", "score4", s4.Caption
SaveSetting "SAON", "highscores", "score5", s5.Caption
SaveSetting "SAON", "highscores", "score6", s6.Caption
SaveSetting "SAON", "highscores", "score7", s7.Caption
SaveSetting "SAON", "highscores", "score8", s8.Caption
SaveSetting "SAON", "highscores", "score9", s9.Caption
SaveSetting "SAON", "highscores", "score10", s10.Caption

End Sub

Sub loader()
c1.Caption = GetSetting("SAON", "highscores", "nick1")
c2.Caption = GetSetting("SAON", "highscores", "nick2")
c3.Caption = GetSetting("SAON", "highscores", "nick3")
c4.Caption = GetSetting("SAON", "highscores", "nick4")
c5.Caption = GetSetting("SAON", "highscores", "nick5")
c6.Caption = GetSetting("SAON", "highscores", "nick6")
c7.Caption = GetSetting("SAON", "highscores", "nick7")
c8.Caption = GetSetting("SAON", "highscores", "nick8")
c9.Caption = GetSetting("SAON", "highscores", "nick9")
c10.Caption = GetSetting("SAON", "highscores", "nick10")

s1.Caption = GetSetting("SAON", "highscores", "score1")
s2.Caption = GetSetting("SAON", "highscores", "score2")
s3.Caption = GetSetting("SAON", "highscores", "score3")
s4.Caption = GetSetting("SAON", "highscores", "score4")
s5.Caption = GetSetting("SAON", "highscores", "score5")
s6.Caption = GetSetting("SAON", "highscores", "score6")
s7.Caption = GetSetting("SAON", "highscores", "score7")
s8.Caption = GetSetting("SAON", "highscores", "score8")
s9.Caption = GetSetting("SAON", "highscores", "score9")
s10.Caption = GetSetting("SAON", "highscores", "score10")
End Sub
