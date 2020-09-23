VERSION 5.00
Begin VB.Form battle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                Enemy Attack"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Start battle"
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox pchino 
      BorderStyle     =   0  'None
      Height          =   1660
      Left            =   150
      Picture         =   "weaplist.frx":0000
      ScaleHeight     =   1665
      ScaleWidth      =   3195
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   3200
   End
   Begin VB.PictureBox chino 
      BorderStyle     =   0  'None
      Height          =   1660
      Left            =   150
      Picture         =   "weaplist.frx":11A02
      ScaleHeight     =   1665
      ScaleWidth      =   3195
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   3200
   End
   Begin VB.PictureBox armcop 
      BorderStyle     =   0  'None
      Height          =   1660
      Left            =   150
      Picture         =   "weaplist.frx":23404
      ScaleHeight     =   1665
      ScaleWidth      =   3195
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   3200
   End
   Begin VB.PictureBox tex 
      BorderStyle     =   0  'None
      Height          =   1660
      Left            =   150
      Picture         =   "weaplist.frx":34E06
      ScaleHeight     =   1665
      ScaleWidth      =   3180
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.PictureBox cop 
      BorderStyle     =   0  'None
      Height          =   1660
      Left            =   240
      Picture         =   "weaplist.frx":46808
      ScaleHeight     =   1665
      ScaleWidth      =   3075
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox joe 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1660
         Left            =   30
         Picture         =   "weaplist.frx":5820A
         ScaleHeight     =   1665
         ScaleWidth      =   3180
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   1800
         Width           =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   3240
         X2              =   0
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label enmaxhealth 
         AutoSize        =   -1  'True
         Caption         =   " <>"
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label enhealth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<>"
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   1800
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   510
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Punch"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< Use"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.ListBox weapons 
      Height          =   1620
      ItemData        =   "weaplist.frx":69C0C
      Left            =   120
      List            =   "weaplist.frx":69C0E
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox attacker 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox weapboost 
      Height          =   285
      Left            =   480
      TabIndex        =   22
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox punchboost 
      Height          =   285
      Left            =   360
      TabIndex        =   23
      Text            =   "1"
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label maxhealth 
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   4800
      Width           =   180
   End
   Begin VB.Label Label6 
      Caption         =   "/"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label yourhealth 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<>"
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Top             =   4800
      Width           =   180
   End
   Begin VB.Label Label4 
      Caption         =   "Your Health:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Weapons in Inventory:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "battle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If weapons.List(weapons.ListIndex) = "Special .12" Then
enhealth.Caption = enhealth.Caption - 50 * weapboost.Text

ElseIf weapons.List(weapons.ListIndex) = "Duece-Duece" Then
enhealth.Caption = enhealth.Caption - 100 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Magnum" Then
enhealth.Caption = enhealth.Caption - 150 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Glock .45" Then
enhealth.Caption = enhealth.Caption - 200 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "12 Gauge" Then
enhealth.Caption = enhealth.Caption - 300 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Assault Rifle" Then
enhealth.Caption = enhealth.Caption - 400 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Rocket Launcher" Then
enhealth.Caption = enhealth.Caption - 600 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Particle Cannon" Then
enhealth.Caption = enhealth.Caption - 1500 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Pulse Cannon" Then
enhealth.Caption = enhealth.Caption - 9000 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Soul Breaker" Then
enhealth.Caption = enhealth.Caption - 12500 * Val(weapboost.Text)

ElseIf weapons.List(weapons.ListIndex) = "Dark Beam" Then
enhealth.Caption = enhealth.Caption - 20000 * Val(weapboost.Text)
End If

If enhealth.Caption = "0" Or Val(enhealth.Caption) < 0 Then GoTo youwon



If attacker.Text = "JOE" Then
yourhealth.Caption = yourhealth.Caption - 100

ElseIf attacker.Text = "TEX" Then
yourhealth.Caption = yourhealth.Caption - 150

ElseIf attacker.Text = "COP" Then
yourhealth.Caption = yourhealth.Caption - 200

ElseIf attacker.Text = "ARMCOP" Then
yourhealth.Caption = yourhealth.Caption - 300

ElseIf attacker.Text = "CHINO" Then
yourhealth.Caption = yourhealth.Caption - 400

ElseIf attacker.Text = "PCHINO" Then
yourhealth.Caption = yourhealth.Caption - 600

End If

If yourhealth.Caption = "0" Then
MsgBox "you have been defeated.", vbExclamation, "You Lost"
main.health.Caption = yourhealth.Caption
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

ElseIf yourhealth.Caption < 0 Then
yourhealth = "0"
MsgBox "You were defeated. And lost half your cash."
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
main.totcash.Caption = main.totcash.Caption / 2
main.health.Caption = yourhealth.Caption
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

ElseIf enhealth.Caption = "0" Then
MsgBox "you defeated Enemy, and stole half their cash."

If attacker.Text = "JOE" Then
main.totcash.Caption = main.totcash.Caption + 200

ElseIf attacker.Text = "TEX" Then
main.totcash.Caption = main.totcash.Caption + 400

ElseIf attacker.Text = "COP" Then
main.totcash.Caption = main.totcash.Caption + 1000

ElseIf attacker.Text = "ARMCOP" Then
main.totcash.Caption = main.totcash.Caption + 2000

ElseIf attacker.Text = "CHINO" Then
main.totcash.Caption = main.totcash.Caption + 15000

ElseIf attacker.Text = "PCHINO" Then
main.totcash.Caption = main.totcash.Caption + 50000
End If

main.health.Caption = yourhealth.Caption
main.inven.Enabled = True
main.locations.Enabled = True
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

ElseIf enhealth.Caption < 0 Then
youwon:
enhealth = "0"
main.health.Caption = yourhealth.Caption
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
MsgBox "you defeated Enemy, and stole half their cash."

If attacker.Text = "JOE" Then
main.totcash.Caption = main.totcash.Caption + 200

ElseIf attacker.Text = "TEX" Then
main.totcash.Caption = main.totcash.Caption + 400

ElseIf attacker.Text = "COP" Then
main.totcash.Caption = main.totcash.Caption + 1000

ElseIf attacker.Text = "ARMCOP" Then
main.totcash.Caption = main.totcash.Caption + 200000

ElseIf attacker.Text = "CHINO" Then
main.totcash.Caption = main.totcash.Caption + 500000

ElseIf attacker.Text = "PCHINO" Then
main.totcash.Caption = main.totcash.Caption + 1000000
End If

joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
End If

End Sub

Private Sub Command2_Click()
main.health.Caption = yourhealth.Caption
MsgBox "You Forfeited, you lost half your cash when you ran.", vbExclamation, "Run"
main.totcash.Caption = Val(main.totcash.Caption) / 2
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
End Sub

Private Sub Command3_Click()
enhealth.Caption = enhealth.Caption - 20 * punchboost.Text

If attacker.Text = "JOE" Then
yourhealth.Caption = yourhealth.Caption - 100

ElseIf attacker.Text = "TEX" Then
yourhealth.Caption = yourhealth.Caption - 150

ElseIf attacker.Text = "COP" Then
yourhealth.Caption = yourhealth.Caption - 200

ElseIf attacker.Text = "ARMCOP" Then
yourhealth.Caption = yourhealth.Caption - 300

ElseIf attacker.Text = "CHINO" Then
yourhealth.Caption = yourhealth.Caption - 400

ElseIf attacker.Text = "PCHINO" Then
yourhealth.Caption = yourhealth.Caption - 600
End If

If yourhealth.Caption = "0" Then
MsgBox "you have been defeated.", vbExclamation, "You Lost"
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
main.totcash.Caption = main.totcash.Caption / 2
main.health.Caption = yourhealth.Caption
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

ElseIf yourhealth.Caption < 0 Then
yourhealth = "0"
MsgBox "you have been defeated.", vbExclamation, "You Lost"
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
main.totcash.Caption = main.totcash.Caption / 2
main.health.Caption = yourhealth.Caption
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

ElseIf enhealth.Caption = "0" Then
MsgBox "you defeated Enemy, and stole half their cash."
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
main.health.Caption = yourhealth.Caption
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

ElseIf enhealth.Caption < 0 Then
enhealth = "0"

main.health.Caption = yourhealth.Caption
MsgBox "you defeated Enemy, and stole half their cash."
main.inven.Enabled = True
main.locations.Enabled = True
tele.Enabled = True
gunshop.Enabled = True
bankman.Enabled = True
sabank.Enabled = True
main.totcash.Caption = main.totcash.Caption + 2000
joe.Visible = False
tex.Visible = False
cop.Visible = False
armcop.Visible = False
chino.Visible = False
pchino.Visible = False
Me.Hide
enhealth.Caption = "<>"
enmaxhealth.Caption = "<>"
yourhealth.Caption = "<>"
maxhealth.Caption = "<>"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True

End If
End Sub

Private Sub Command4_Click()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
attack
yourhealth.Caption = main.health.Caption
maxhealth.Caption = main.mhealth.Caption
If weapons.ListCount = "0" Then
Command1.Enabled = False
End If
Command4.Enabled = False
End Sub

Private Sub Form_Load()
maxhealth.Caption = main.mhealth
yourhealth.Caption = main.health
End Sub

Sub attack()
If attacker.Text = "JOE" Then
enmaxhealth.Caption = main.joe.Text
joe.Visible = True

ElseIf attacker.Text = "TEX" Then
enmaxhealth.Caption = main.tex.Text
tex.Visible = True

ElseIf attacker.Text = "COP" Then
enmaxhealth.Caption = main.cop.Text
cop.Visible = True

ElseIf attacker.Text = "ARMCOP" Then
enmaxhealth.Caption = main.armcop.Text
armcop.Visible = True

ElseIf attacker.Text = "CHINO" Then
enmaxhealth.Caption = main.chino.Text
chino.Visible = True

ElseIf attacker.Text = "PCHINO" Then
enmaxhealth.Caption = main.pchino.Text
pchino.Visible = True

End If
enhealth.Caption = enmaxhealth.Caption
End Sub
