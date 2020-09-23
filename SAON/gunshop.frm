VERSION 5.00
Begin VB.Form gunshop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terminal G's Weapons Store"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command16 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   51
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   49
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   47
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   45
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   43
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   41
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   37
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   36
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   34
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   32
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Exit GunShop"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   29
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   8880
      TabIndex        =   27
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   25
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   23
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Purchase"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   375
      Left            =   30
      TabIndex        =   4
      Top             =   5160
      Width           =   2385
      Begin VB.Label totsave 
         AutoSize        =   -1  'True
         Caption         =   "Total Savings:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weapon Information"
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   4580
      Width           =   3735
      Begin VB.Label Label6 
         Caption         =   "Damage:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cost:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label damage 
         Caption         =   "<>"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label cost 
         Caption         =   "<>"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.ListBox yourweaps 
      Height          =   2595
      ItemData        =   "gunshop.frx":0000
      Left            =   120
      List            =   "gunshop.frx":0007
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FF8080&
      Caption         =   " Health Splitter"
      Height          =   255
      Left            =   6480
      TabIndex        =   50
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label23 
      BackColor       =   &H008080FF&
      Caption         =   " Soul Breaker"
      Height          =   255
      Left            =   2640
      TabIndex        =   48
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FF8080&
      Caption         =   " Time Warp"
      Height          =   255
      Left            =   6480
      TabIndex        =   46
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label22 
      BackColor       =   &H008080FF&
      Caption         =   " Dark Beam"
      Height          =   255
      Left            =   2640
      TabIndex        =   44
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label21 
      BackColor       =   &H0080C0FF&
      Caption         =   " Tesla Taser"
      Height          =   255
      Left            =   2640
      TabIndex        =   42
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label20 
      BackColor       =   &H008080FF&
      Caption         =   " 9000 Mega Joule Pulse Cannon "
      Height          =   255
      Left            =   2640
      TabIndex        =   40
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FF8080&
      Caption         =   " Weapon Booster"
      Height          =   255
      Left            =   6480
      TabIndex        =   39
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Terminals Items:"
      Height          =   195
      Index           =   2
      Left            =   6480
      TabIndex        =   38
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FF8080&
      Caption         =   " Hyper Punch"
      Height          =   255
      Left            =   6480
      TabIndex        =   35
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF8080&
      Caption         =   " Sheild"
      Height          =   255
      Left            =   6480
      TabIndex        =   33
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   5640
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF8080&
      Caption         =   " BackPack"
      Height          =   255
      Left            =   6480
      TabIndex        =   31
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      Caption         =   " Medi-Kit"
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   " Steroids"
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080C0FF&
      Caption         =   " Particle Beam Cannon (ex lease)"
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      Caption         =   " Rocket Launcher"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      Caption         =   " Serbian Assault Rifle .50 cal."
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      Caption         =   " 12 gauge sawn off Shotgun"
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Caption         =   " Glock .45 cal."
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   " Smith and Western Magnum"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   " Duece-Duece 22 cal."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   " Special 12 calibur"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Your Weapons/Items:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   2480
      X2              =   2480
      Y1              =   0
      Y2              =   5640
   End
   Begin VB.Label Label1 
      Caption         =   "Terminals Weapons:"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "gunshop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
If Val(main.totcash) > 250 Then
main.totcash = main.totcash - 250
battle.weapons.AddItem "Special .12"
yourweaps.AddItem "special .12"
update
Command1.Enabled = False
Else
msg
End If
End Sub
Sub msg()
MsgBox "Not Enough Cash To purchase", vbExclamation, "Funds"
End Sub
Private Sub Command10_Click()
If Val(main.totcash) > 25000 Then
main.totcash = main.totcash - 25000
main.medikits.Text = main.medikits.Text + 1
main.medi.Caption = "Medi-Kit " & main.medikits.Text
update
Else
msg
End If
End Sub

Private Sub Command11_Click()
Me.Hide
End Sub

Private Sub Command12_Click()
If Val(main.totcash) > 50000 Then
main.totcash = main.totcash - 50000
main.maxroom = Val(main.maxroom) + 50
update
Else
msg
End If
End Sub

Private Sub Command13_Click()
If Val(main.totcash) > 100000 Then
main.totcash = main.totcash - 100000
main.mhealth.Caption = Val(main.mhealth.Caption + 100)
update
Else
msg
End If
End Sub

Private Sub Command14_Click()
If Val(main.totcash) > 2000000 Then
main.totcash = main.totcash - 2000000
battle.punchboost.Text = Val(battle.punchboost.Text) + 1
update
Else
msg
End If
End Sub

Private Sub Command15_Click()
If Val(main.totcash) > 2000000 Then
main.totcash = main.totcash - 2000000
battle.weapboost.Text = Val(battle.weapboost.Text) + 1
update
Else
msg
End If
End Sub

Private Sub Command16_Click()
If Val(main.totcash) > 2000000 Then
main.totcash = main.totcash - 2000000
main.joe.Text = Val(main.joe.Text) / 2
main.tex.Text = Val(main.tex.Text) / 2
main.cop.Text = Val(main.cop.Text) / 2
main.armcop.Text = Val(main.armcop.Text) / 2
main.chino.Text = Val(main.chino.Text) / 2
update
MsgBox "All enemies health have been halved.", vbExclamation, "Health Splitter"
Else
msg
End If
End Sub

Private Sub Command17_Click()
If Val(main.totcash) > 2000000 Then
main.totcash = main.totcash - 2000000
main.days.Caption = Val(main.days.Caption) + 5
update
Else
msg
End If
End Sub

Private Sub Command18_Click()
If Val(main.totcash) > 100000 Then
Command18.Enabled = False
main.totcash = main.totcash - 100000
battle.weapons.AddItem "Tesla"
yourweaps.AddItem "Tesla"
update
Else
msg
End If
End Sub

Private Sub Command19_Click()
If Val(main.totcash) > 500000 Then
Command19.Enabled = False
main.totcash = main.totcash - 500000
battle.weapons.AddItem "Pusle Cannon"
yourweaps.AddItem "Pulse Cannon"
update
Command19.Enabled = False
Else
msg
End If
End Sub

Private Sub Command2_Click()

If Val(main.totcash) > 4350 Then
Command2.Enabled = False
main.totcash = main.totcash - 4350
battle.weapons.AddItem "Duece-Duece"
yourweaps.AddItem "Duece-Duece"
update
Else
msg
End If
End Sub

Private Sub Command20_Click()
If Val(main.totcash) > 1000000 Then
Command7.Enabled = False
main.totcash = main.totcash - 1000000
battle.weapons.AddItem "Dark Beam"
yourweaps.AddItem "Dark Beam"
update
Command20.Enabled = False
Else
msg
End If
End Sub

Private Sub Command21_Click()
If Val(main.totcash) > 750000 Then
Command21.Enabled = False
main.totcash = main.totcash - 750000
battle.weapons.AddItem "Soul Breaker"
yourweaps.AddItem "Soul Breaker"
update
Command21.Enabled = False
Else
msg
End If
End Sub

Private Sub Command3_Click()

If Val(main.totcash) > 9700 Then
Command3.Enabled = False
main.totcash = main.totcash - 9700
battle.weapons.AddItem "Magnum"
yourweaps.AddItem "Magnum"
update
Else
msg
End If
End Sub

Private Sub Command4_Click()
If Val(main.totcash) > 13250 Then
Command4.Enabled = False
main.totcash = main.totcash - 13250
battle.weapons.AddItem "Glock .45"
yourweaps.AddItem "Glock .45"
update
Else
msg
End If
End Sub

Private Sub Command5_Click()

If Val(main.totcash) > 21500 Then
Command5.Enabled = False
main.totcash = main.totcash - 21500
battle.weapons.AddItem "12 Gauge"
yourweaps.AddItem "12 Gauge"
update
Else
msg
End If
End Sub

Private Sub Command6_Click()

If Val(main.totcash) > 30000 Then
Command6.Enabled = False
main.totcash = main.totcash - 30000
battle.weapons.AddItem "Assault Rifle"
yourweaps.AddItem "Assault Rifle"
update
Else
msg
End If
End Sub

Private Sub Command7_Click()

If Val(main.totcash) > 45000 Then
Command7.Enabled = False
main.totcash = main.totcash - 45000
battle.weapons.AddItem "Rocket Launcher"
yourweaps.AddItem "Rocket launcher"
update
Else
msg
End If
End Sub

Private Sub Command8_Click()

If Val(main.totcash) > 130000 Then
Command8.Enabled = False
main.totcash = main.totcash - 130000
battle.weapons.AddItem "Particle Cannon"
yourweaps.AddItem "Particle Cannon"
update
Else
msg
End If

End Sub

Private Sub Command9_Click()
If Val(main.totcash) > 25000 Then
main.totcash = main.totcash - 25000
main.steroid.Text = main.steroid.Text + 1
main.ster.Caption = "Steroids " & main.steroid.Text
update
Else
msg
End If
End Sub

Private Sub Form_Load()
yourweaps.Clear
totsave.Caption = "Total Savings: " & main.totcash.Caption
End Sub

Sub update()
totsave.Caption = "Total Savings: " & main.totcash.Caption
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
update
End Sub

Private Sub Form_Resize()
update
End Sub

Private Sub Label10_Click()
cost.Caption = "$30000"
damage.Caption = "400"
End Sub

Private Sub Label11_Click()
cost.Caption = "$45000"
damage.Caption = "600"
End Sub

Private Sub Label12_Click()
cost.Caption = "$130000"
damage.Caption = "1500"
End Sub

Private Sub Label13_Click()
cost.Caption = "$2500"
damage.Caption = "Full health Restore"
End Sub

Private Sub Label14_Click()
cost.Caption = "$50000"
damage.Caption = "Increases Drug Storage Capacity."
End Sub

Private Sub Label15_Click()
cost.Caption = "$100000"
damage.Caption = "Doubles Max Health."
End Sub

Private Sub Label16_Click()
cost.Caption = "25000"
damage.Caption = "Doubles punch intensity."
End Sub

Private Sub Label17_Click()
cost.Caption = "300000"
damage.Caption = "Cuts all enemies health in half."
End Sub

Private Sub Label18_Click()
cost.Caption = "5000000"
damage.Caption = "Doubles any weapons fire power."
End Sub

Private Sub Label19_Click()
cost.Caption = "$300000"
damage.Caption = "Takes you back in time 5 days."
End Sub

Private Sub Label2_Click()
cost.Caption = "$250"
damage.Caption = "50"
End Sub

Private Sub Label20_Click()
cost.Caption = "$500000"
damage.Caption = "9000"
End Sub

Private Sub Label21_Click()
cost.Caption = "$100000"
damage.Caption = "1000"
End Sub

Private Sub Label22_Click()
cost.Caption = "$1000000"
damage.Caption = "20000"
End Sub


Private Sub Label23_Click()
cost.Caption = "$750000"
damage.Caption = "12500"
End Sub

Private Sub Label3_Click()
cost.Caption = "$4350"
damage.Caption = "100"
End Sub

Private Sub Label4_Click()
cost.Caption = "$25000"
damage.Caption = "Increase max health by 20."
End Sub

Private Sub Label7_Click()
cost.Caption = "$9700"
damage.Caption = "150"
End Sub

Private Sub Label8_Click()
cost.Caption = "$13250"
damage.Caption = "200"
End Sub

Private Sub Label9_Click()
cost.Caption = "21500"
damage.Caption = "300"
End Sub
