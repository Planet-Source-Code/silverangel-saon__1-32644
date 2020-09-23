VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SA Organised Narcosis 2002 "
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   ForeColor       =   &H00E0E0E0&
   Icon            =   "splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   20
      TabIndex        =   3
      Top             =   4800
      Width           =   4850
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   120
         Width           =   90
      End
      Begin VB.Label highestscore 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Highest Score:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1050
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox nick 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   840
      MaxLength       =   12
      TabIndex        =   1
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "splash.frx":0442
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Alias Below:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SaveSetting "SAON", "splash", "nick", nick.Text
main.nickname.Caption = nick.Text
Unload Me
main.Show
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub

Sub Form_Load()
highestscore.Caption = GetSetting("SAON", "highscores", "score1")
nick.Text = GetSetting("SAON", "splash", "nick")
If highestscore.Caption = "" Then
highestscore.Caption = "Yet to be set."
End If
If nick.Text = "" Then
nick.Text = "Dealer"
End If
End Sub


Private Sub Text1_GotFocus()
Command1.SetFocus
End Sub

