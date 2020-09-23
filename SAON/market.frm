VERSION 5.00
Begin VB.Form silverclub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   Silvers Nightclub"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Access Computer Terminal"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nightclub Patrons"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.ListBox patrons 
         Height          =   3180
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "silverclub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
patrons.Clear
patrons.AddItem "Veronica"
patrons.AddItem "Mark"
patrons.AddItem "James"
patrons.AddItem "Gail"
patrons.AddItem "Rebecca"
patrons.AddItem "Thomas"
patrons.AddItem "Ivory"
patrons.AddItem "Trey"
patrons.AddItem "Sally"
patrons.AddItem "Kirk"
patrons.AddItem "Dunkirk"
patrons.AddItem "Redhead"
patrons.AddItem "Super Nerd"
End Sub
