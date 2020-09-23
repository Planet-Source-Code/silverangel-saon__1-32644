VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Organized Narcosis 2002, 1.1"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   ForeColor       =   &H00E0E0E0&
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   240
      Picture         =   "about.frx":0442
      ScaleHeight     =   720
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label4 
      Caption         =   "Check the Website for Updates:"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "http://home.dencity.com/narcosis.htm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Graphics by: C-MP"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Coded by: SilverAngel"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label3_Click()
MsgBox "Check out our website for the latest updates and other SA C-MP software.", vbInformation, "SA C-MP"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.MousePointer = vbUpArrow
End Sub
