VERSION 5.00
Begin VB.Form bankman 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pay Debt"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label totowe 
      AutoSize        =   -1  'True
      Caption         =   "Total Savings: "
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2745
   End
   Begin VB.Label totsave 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "bankman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Val(main.totcash.Caption) < Val(Text1.Text) Then
MsgBox "Not Enough Funds to Pay off Debt.", vbExclamation, "Invalid Funds"
Else
main.totcash.Caption = main.totcash.Caption - Text1.Text
main.totdebt.Caption = main.totdebt.Caption - Text1.Text
totsave.Caption = "Total Savings: " & main.totcash.Caption
totowe.Caption = "Total Debt: " & main.totdebt.Caption
main.debtcheck.Text = "paid"
End If
If main.totdebt.Caption < 0 Then
main.totdebt.Caption = "0"
totsave.Caption = "Total Savings: " & main.totcash.Caption
totowe.Caption = "Total Debt: " & main.totdebt.Caption
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
totsave.Caption = "Total Savings: " & main.totcash.Caption
totowe.Caption = "Total Debt: " & main.totdebt.Caption
End Sub
