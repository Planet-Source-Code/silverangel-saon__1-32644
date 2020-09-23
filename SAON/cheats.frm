VERSION 5.00
Begin VB.Form cheats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheat Menu"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3345
   ForeColor       =   &H00E0E0E0&
   Icon            =   "cheats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cheat!"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox cheat 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "SA Cheat Menu Enter Cheats Below."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "cheats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cheat.Text = "MILLISHEILD" Then
main.cheatfixer.Text = "cheated"
main.mhealth.Caption = Val(main.mhealth.Caption * 2)
MsgBox "Health Doubled.", vbExclamation, "Cheater"

ElseIf cheat.Text = "RIOSREVENGE" Then
main.cheatfixer.Text = "cheated!"
battle.weapons.Clear
gunshop.yourweaps.Clear
battle.weapons.AddItem "Special .12"
battle.weapons.AddItem "Duece Duece"
battle.weapons.AddItem "Magnum"
battle.weapons.AddItem "Glock .45"
battle.weapons.AddItem "12 Gauge"
battle.weapons.AddItem "Assault Rifle"
battle.weapons.AddItem "Rocket Launcher"
battle.weapons.AddItem "Tesla Taser"
battle.weapons.AddItem "Particle Cannon"
battle.weapons.AddItem "Pulse Cannon"
battle.weapons.AddItem "Soul Breaker"
battle.weapons.AddItem "Dark Beam"

gunshop.yourweaps.AddItem "Special .12"
gunshop.yourweaps.AddItem "Duece Duece"
gunshop.yourweaps.AddItem "Magnum"
gunshop.yourweaps.AddItem "Glock .45"
gunshop.yourweaps.AddItem "12 Gauge"
gunshop.yourweaps.AddItem "Assault Rifle"
gunshop.yourweaps.AddItem "Rocket Launcher"
gunshop.yourweaps.AddItem "Tesla Taser"
gunshop.yourweaps.AddItem "Particle Cannon"
gunshop.yourweaps.AddItem "Pulse Cannon"
gunshop.yourweaps.AddItem "Soul Breaker"
gunshop.yourweaps.AddItem "Dark Beam"

gunshop.Command1.Enabled = False
gunshop.Command2.Enabled = False
gunshop.Command3.Enabled = False
gunshop.Command4.Enabled = False
gunshop.Command5.Enabled = False
gunshop.Command6.Enabled = False
gunshop.Command7.Enabled = False
gunshop.Command18.Enabled = False
gunshop.Command8.Enabled = False
gunshop.Command19.Enabled = False
gunshop.Command21.Enabled = False
gunshop.Command20.Enabled = False
MsgBox "You now have All Weapons.", vbExclamation, "Cheater"

ElseIf cheat.Text = "INEEDSOMEDOUGH" Then
main.cheatfixer.Text = "cheated!"
main.totcash.Caption = Val(main.totcash.Caption * 10)
MsgBox "Your Money has been multiplied 10 times.", vbExclamation, "Cheater"

ElseIf cheat.Text = "PUMPMEUP" Then
main.cheatfixer.Text = "cheated!"
main.steroid.Text = Val(main.steroid.Text + 500)
main.medikits.Text = Val(main.medikits.Text + 500)
main.ster.Caption = "Steroids " & main.steroid.Text
main.medi.Caption = "Medi-Kits " & main.medikits.Text
MsgBox "You received 500 Steroids and 500 Medikits.", vbExclamation, "Cheater"

ElseIf cheat.Text = "WARPME" Then
main.cheatfixer.Text = "cheated!"
main.days.Caption = main.maxdays.Caption
MsgBox "You have gone back in time to day one.", vbExclamation, "Time Warp"
Else
cheater
End If
cheat.Text = ""
End Sub

Sub cheater()
MsgBox "Thats not a cheat, please try again.", vbExclamation, "Cheats"
End Sub

