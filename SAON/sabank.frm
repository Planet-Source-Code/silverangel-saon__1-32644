VERSION 5.00
Begin VB.Form sabank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Narcosis Federal Bank"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Your Account"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Deposit"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox depamount 
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Text            =   "0"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Withdraw"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox withamount 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label handcash 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label7 
         Caption         =   "$"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "  $"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1590
         Width           =   180
      End
      Begin VB.Label Label5 
         Caption         =   "  $"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "$ on hand:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label storecash 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Savings:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect To bank"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label stat 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "**Not Connected to Narcosis Bank...**"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2800
   End
End
Attribute VB_Name = "sabank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errhan
If Val(withamount.Text) > Val(storecash.Caption) Then
MsgBox "Invalid funds to complete Transaction.", vbExclamation, "Not Enough Funds"
Else
storecash.Caption = Val(storecash.Caption) - Val(withamount.Text)
main.totcash.Caption = Val(main.totcash.Caption) + Val(withamount.Text)
handcash.Caption = Val(handcash.Caption) + Val(withamount.Text)
withamount.Text = "0"
End If
Exit Sub
errhan:
MsgBox "Invalid amount to withdraw.", vbExclamation, "Withdraw"
End Sub

Private Sub Command2_Click()
On Error GoTo errhan
If Val(depamount.Text) > Val(handcash.Caption) Then
MsgBox "Invalid funds to complete Transaction.", vbExclamation, "Not Enough Funds"
Else
handcash.Caption = Val(handcash.Caption - depamount.Text)
main.totcash.Caption = Val(main.totcash.Caption) - Val(depamount.Text)
storecash.Caption = Val(storecash.Caption) + Val(depamount.Text)
depamount.Text = "0"
End If
Exit Sub
errhan:
MsgBox "Invalid Amount To deposit.", vbExclamation, "Deposit"
End Sub

Private Sub Command3_Click()
Command3.Visible = False
Frame1.Visible = True
handcash.Caption = main.totcash.Caption
stat.Caption = " Connected to SA bank..."
depamount.Text = handcash.Caption
withamount.Text = storecash.Caption
End Sub

Private Sub Command4_Click()
Me.Hide
Command3.Visible = True
stat.Caption = "**Not Connected to Narcosis Bank...**"
Frame1.Visible = False
End Sub

Private Sub depamount_Change()
If Left(depamount.Text, 1) = "-" Then
depamount.Text = Right(depamount.Text, Len(depamount.Text) - 1)
End If

If Mid(depamount.Text, 2, 1) = "." Then
depamount.Text = Left(depamount.Text, 1)
Command2.SetFocus

ElseIf Mid(depamount.Text, 3, 1) = "." Then
depamount.Text = Left(depamount.Text, 2)
Command2.SetFocus

ElseIf Mid(depamount.Text, 4, 1) = "." Then
depamount.Text = Left(depamount.Text, 3)
Command2.SetFocus

ElseIf Mid(depamount.Text, 5, 1) = "." Then
depamount.Text = Left(depamount.Text, 4)
Command2.SetFocus

ElseIf Mid(depamount.Text, 6, 1) = "." Then
depamount.Text = Left(depamount.Text, 5)
Command2.SetFocus

ElseIf Mid(depamount.Text, 7, 1) = "." Then
depamount.Text = Left(depamount.Text, 6)
Command2.SetFocus

ElseIf Mid(depamount.Text, 8, 1) = "." Then
depamount = Left(depamount.Text, 7)
Command2.SetFocus

ElseIf Mid(depamount.Text, 9, 1) = "." Then
depamount.Text = Left(depamount.Text, 8)
Command2.SetFocus

ElseIf Mid(depamount.Text, 10, 1) = "." Then
depamount.Text = Left(depamount.Text, 9)
Command2.SetFocus

ElseIf Mid(depamount.Text, 11, 1) = "." Then
depamount.Text = Left(depamount.Text, 10)
Command2.SetFocus

ElseIf Mid(depamount.Text, 12, 1) = "." Then
depamount.Text = Left(depamount.Text, 11)
Command2.SetFocus

ElseIf Mid(depamount.Text, 13, 1) = "." Then
depamount.Text = Left(depamount.Text, 12)
Command2.SetFocus

ElseIf Mid(depamount.Text, 14, 1) = "." Then
depamount.Text = Left(depamount.Text, 13)
Command2.SetFocus

ElseIf Mid(depamount.Text, 15, 1) = "." Then
depamount.Text = Left(depamount.Text, 14)
Command2.SetFocus

End If

End Sub



Private Sub withamount_Change()
If Left(withamount.Text, 1) = "-" Then
withamount.Text = Right(withamount.Text, Len(withamount.Text) - 1)
End If

If Mid(withamount.Text, 2, 1) = "." Then
withamount.Text = Left(withamount.Text, 1)
Command1.SetFocus

ElseIf Mid(withamount.Text, 3, 1) = "." Then
withamount.Text = Left(withamount.Text, 2)
Command1.SetFocus

ElseIf Mid(withamount.Text, 4, 1) = "." Then
withamount.Text = Left(withamount.Text, 3)
Command1.SetFocus

ElseIf Mid(withamount.Text, 5, 1) = "." Then
withamount.Text = Left(withamount.Text, 4)
Command1.SetFocus

ElseIf Mid(withamount.Text, 6, 1) = "." Then
withamount.Text = Left(withamount.Text, 5)
Command1.SetFocus

ElseIf Mid(withamount.Text, 7, 1) = "." Then
withamount.Text = Left(withamount.Text, 6)
Command1.SetFocus

ElseIf Mid(withamount.Text, 8, 1) = "." Then
withamount = Left(withamount.Text, 7)
Command1.SetFocus

ElseIf Mid(withamount.Text, 9, 1) = "." Then
withamount.Text = Left(withamount.Text, 8)
Command1.SetFocus

ElseIf Mid(withamount.Text, 10, 1) = "." Then
withamount.Text = Left(withamount.Text, 9)
Command1.SetFocus

ElseIf Mid(withamount.Text, 11, 1) = "." Then
withamount.Text = Left(withamount.Text, 10)
Command1.SetFocus

ElseIf Mid(withamount.Text, 12, 1) = "." Then
withamount.Text = Left(withamount.Text, 11)
Command1.SetFocus

ElseIf Mid(withamount.Text, 13, 1) = "." Then
withamount.Text = Left(withamount.Text, 12)
Command1.SetFocus

ElseIf Mid(withamount.Text, 14, 1) = "." Then
withamount.Text = Left(withamount.Text, 13)
Command1.SetFocus

ElseIf Mid(withamount.Text, 15, 1) = "." Then
withamount.Text = Left(withamount.Text, 14)
Command1.SetFocus

End If
End Sub
