VERSION 5.00
Begin VB.Form tele 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telephone"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1815
   Icon            =   "tele.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   1815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Speed Dial"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command6 
         Caption         =   "Chino"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Armour Cop"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Chinos Wife"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cop"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Tex"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Joe"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "tele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
main.inven.Enabled = False
main.locations.Enabled = False
MsgBox "You Contacted Joe, and organised a Fight. the winner gets $500.", vbExclamation, "Battle"
battle.attacker.Text = "JOE"
battle.Show
End Sub

Private Sub Command2_Click()
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
main.inven.Enabled = False
main.locations.Enabled = False
MsgBox "You Contacted Tex, and organised a Fight. the winner gets $800.", vbExclamation, "Battle"
battle.attacker.Text = "TEX"
battle.Show
End Sub

Private Sub Command3_Click()
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
main.inven.Enabled = False
main.locations.Enabled = False
MsgBox "You Contacted The cops, and said you need assistance. the police sent out a squad car and you will battle when they arrive.", vbExclamation, "Battle"
battle.attacker.Text = "COP"
battle.Show
End Sub

Private Sub Command4_Click()
MsgBox "You rendezvous with Chinos wife, it'll cost $500 for a blowjob that will fully restore your health."
If Val(main.totcash) > 500 Then
main.health.Caption = main.mhealth.Caption
main.pchino.Text = main.pchino.Text + 200
main.totcash.Caption = main.totcash.Caption - 500
Else
MsgBox "Sorry Honey, you don't have enough cash"
End If
End Sub

Private Sub Command5_Click()
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
MsgBox "You Contacted Armoured Cop, and claimed to have robbed a bank, upon jetpacking to your location you will battle. ", vbExclamation, "Battle"
battle.attacker.Text = "ARMCOP"
battle.Show
End Sub

Private Sub Command6_Click()
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
MsgBox "You wanna smoke up with Chino?"
battle.attacker.Text = "CHINO"
battle.Show
End Sub


