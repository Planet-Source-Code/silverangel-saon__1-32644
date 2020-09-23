VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SA Organized Narcosis 2002"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox buyamount 
      Height          =   285
      Left            =   3400
      TabIndex        =   109
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox sellamount 
      Height          =   285
      Left            =   3400
      TabIndex        =   108
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame9 
      Height          =   375
      Left            =   7150
      TabIndex        =   104
      Top             =   6120
      Width           =   930
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "About"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   105
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Drugs Available"
      Height          =   3040
      Left            =   250
      TabIndex        =   70
      Top             =   2880
      Width           =   2895
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         TabIndex        =   116
         Top             =   0
         Width           =   420
         Begin VB.Label Label9 
            Caption         =   "Price"
            Height          =   255
            Left            =   20
            TabIndex        =   117
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Shape a11 
         Height          =   225
         Left            =   120
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Shape a10 
         Height          =   225
         Left            =   120
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Shape a9 
         Height          =   225
         Left            =   120
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Shape a8 
         Height          =   225
         Left            =   120
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Shape a7 
         Height          =   225
         Left            =   120
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Shape a6 
         Height          =   225
         Left            =   120
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Shape a5 
         Height          =   225
         Left            =   120
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Shape a4 
         Height          =   225
         Left            =   120
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Shape a3 
         Height          =   225
         Left            =   120
         Top             =   840
         Width           =   2655
      End
      Begin VB.Shape a2 
         Height          =   225
         Left            =   120
         Top             =   600
         Width           =   2655
      End
      Begin VB.Shape a1 
         Height          =   225
         Left            =   120
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   10
         Left            =   1680
         TabIndex        =   103
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   102
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   101
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   100
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   99
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   98
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   97
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   96
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   95
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   94
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   93
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label45 
         Caption         =   "      Cocaine:"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label44 
         Caption         =   "        Heroin:"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label43 
         Caption         =   "     Crystal Meth:"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label42 
         Caption         =   "             IcE:"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label41 
         Caption         =   "           Weed:"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label40 
         Caption         =   "           Ganja:"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "      Angel Dust:"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label38 
         Caption         =   "           Crack:"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label37 
         Caption         =   "          Smack:"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label36 
         Caption         =   "            XTC:"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   1560
         X2              =   1560
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "          Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label pcoc 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   81
         Top             =   360
         Width           =   90
      End
      Begin VB.Label pher 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   80
         Top             =   600
         Width           =   90
      End
      Begin VB.Label pmeth 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   79
         Top             =   840
         Width           =   90
      End
      Begin VB.Label pdust 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   78
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label pweed 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   77
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label pgan 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   76
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label psma 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   75
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label pcra 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   74
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label pxtc 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   73
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label pice 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   72
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label pspe 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   71
         Top             =   2760
         Width           =   90
      End
   End
   Begin VB.TextBox selldeal 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   960
      TabIndex        =   68
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      Caption         =   "Drugs in Inventory"
      Height          =   3040
      Left            =   4920
      TabIndex        =   45
      Top             =   2880
      Width           =   2895
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         TabIndex        =   106
         Top             =   0
         Width           =   615
         Begin VB.Label Label25 
            Caption         =   "Quantity"
            Height          =   255
            Left            =   0
            TabIndex        =   107
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   1560
         Y1              =   120
         Y2              =   3080
      End
      Begin VB.Shape b11 
         Height          =   225
         Left            =   120
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Shape b10 
         Height          =   225
         Left            =   120
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Shape b9 
         Height          =   225
         Left            =   120
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Shape b8 
         Height          =   225
         Left            =   120
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Shape b7 
         Height          =   225
         Left            =   120
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Shape b6 
         Height          =   225
         Left            =   120
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Shape b5 
         Height          =   225
         Left            =   120
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Shape b4 
         Height          =   225
         Left            =   120
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Shape b3 
         Height          =   225
         Left            =   120
         Top             =   840
         Width           =   2655
      End
      Begin VB.Shape b2 
         Height          =   225
         Left            =   120
         Top             =   600
         Width           =   2655
      End
      Begin VB.Shape b1 
         Height          =   225
         Left            =   120
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label speed 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   67
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label ice 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   66
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label xtc 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   65
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label crack 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   64
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label smack 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   63
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label ganja 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   62
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label weed 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   61
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label dust 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   60
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label meth 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   59
         Top             =   840
         Width           =   975
      End
      Begin VB.Label heroin 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   58
         Top             =   600
         Width           =   975
      End
      Begin VB.Label cocaine 
         Caption         =   "0"
         Height          =   195
         Left            =   1800
         TabIndex        =   57
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "      Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label Label22 
         Caption         =   "        XTC:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "      Smack:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "       Crack:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "     Angel Dust:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "      Ganja:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "      Weed:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "         IcE:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "    Crystal Meth:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "      Heroin:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "    Cocaine:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame inven 
      Caption         =   "Bank / Inventory"
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3735
      Begin VB.CommandButton Command13 
         Caption         =   "Bank"
         Height          =   375
         Left            =   1320
         TabIndex        =   121
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Phone"
         Height          =   375
         Left            =   2520
         TabIndex        =   44
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton ster 
         Caption         =   "Steroids 0"
         Height          =   375
         Left            =   2400
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton medi 
         Caption         =   "Medi Kit 0"
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton weap 
         Caption         =   "Gun Shop"
         Height          =   375
         Left            =   2520
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Pay Debt"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label totdebt 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "5000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   405
         TabIndex        =   27
         Top             =   1365
         Width           =   420
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " $"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label totcash 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "3000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   375
         TabIndex        =   25
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " $"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Debt:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cash:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<== Sell"
      Height          =   390
      Left            =   3400
      TabIndex        =   7
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buy ==>"
      Height          =   375
      Left            =   3400
      TabIndex        =   6
      Top             =   3920
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Height          =   375
      Left            =   4890
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   1730
         TabIndex        =   120
         Top             =   120
         Width           =   75
      End
      Begin VB.Label days 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100"
         Height          =   195
         Left            =   1440
         TabIndex        =   30
         Top             =   120
         Width           =   270
      End
      Begin VB.Label maxdays 
         AutoSize        =   -1  'True
         Caption         =   "100"
         Height          =   195
         Left            =   1800
         TabIndex        =   20
         Top             =   120
         Width           =   270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Days Remaining:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      Height          =   375
      Left            =   2510
      TabIndex        =   4
      Top             =   6120
      Width           =   2415
      Begin VB.Label mhealth 
         AutoSize        =   -1  'True
         Caption         =   "500"
         Height          =   195
         Left            =   1560
         TabIndex        =   42
         Top             =   120
         Width           =   270
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "/"
         Height          =   195
         Left            =   1440
         TabIndex        =   41
         Top             =   120
         Width           =   75
      End
      Begin VB.Label health 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "500"
         Height          =   195
         Left            =   1110
         TabIndex        =   21
         Top             =   120
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Height          =   375
      Left            =   20
      TabIndex        =   3
      Top             =   6120
      Width           =   2520
      Begin VB.Label nickname 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   960
         TabIndex        =   17
         Top             =   120
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nickname:"
         Height          =   195
         Left            =   40
         TabIndex        =   16
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame locations 
      Caption         =   "Locations"
      Height          =   2415
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.CommandButton Command10 
         Caption         =   "Columbia"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Jamaica"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Amsterdam"
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Pakistan"
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "London"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Berlin"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Paris"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Belgrade"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame items 
      Height          =   975
      Left            =   2760
      TabIndex        =   38
      Top             =   360
      Width           =   855
      Begin VB.TextBox steroid 
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox medikits 
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Text            =   "0"
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.TextBox buydeal 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5520
      TabIndex        =   69
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame10 
      Height          =   1095
      Left            =   3300
      TabIndex        =   110
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame11 
      Height          =   1095
      Left            =   3300
      TabIndex        =   111
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame enemies 
      Height          =   495
      Left            =   4800
      TabIndex        =   31
      Top             =   720
      Width           =   2775
      Begin VB.TextBox catalyst 
         Height          =   285
         Left            =   480
         TabIndex        =   115
         Text            =   "100000"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox pchino 
         Height          =   285
         Left            =   480
         TabIndex        =   43
         Text            =   "100000"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox chino 
         Height          =   285
         Left            =   480
         TabIndex        =   36
         Text            =   "30000"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox armcop 
         Height          =   285
         Left            =   480
         TabIndex        =   35
         Text            =   "10000"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox cop 
         Height          =   285
         Left            =   480
         TabIndex        =   34
         Text            =   "5000"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox tex 
         Height          =   285
         Left            =   480
         TabIndex        =   33
         Text            =   "400"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox joe 
         Height          =   285
         Left            =   480
         TabIndex        =   32
         Text            =   "200"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox secret 
      Height          =   285
      Left            =   840
      TabIndex        =   118
      Text            =   "99999999999999999"
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox cheatfixer 
      Height          =   285
      Left            =   3120
      TabIndex        =   119
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox debtcheck 
      Height          =   285
      Left            =   5880
      TabIndex        =   122
      Text            =   "unpaid"
      Top             =   1320
      Width           =   150
   End
   Begin VB.Label Label26 
      Caption         =   "/"
      Height          =   255
      Left            =   3840
      TabIndex        =   114
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label room 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   3720
      TabIndex        =   113
      Top             =   3120
      Width           =   90
   End
   Begin VB.Label maxroom 
      AutoSize        =   -1  'True
      Caption         =   "250"
      Height          =   195
      Left            =   3960
      TabIndex        =   112
      Top             =   3120
      Width           =   270
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub buyfix()
MsgBox "Not Enough Money to purchase drug.", vbExclamation, "Lack of funds"
End Sub
Sub roomfix()
MsgBox "Not enough Room, buy more backpacks.", vbExclamation, "No Space"
End Sub

Private Sub buyamount_Change()
If Left(buyamount.Text, 1) = "-" Then
buyamount.Text = Right(buyamount.Text, Len(buyamount.Text) - 1)
End If

If Mid(buyamount.Text, 2, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 1)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 3, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 2)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 4, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 3)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 5, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 4)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 6, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 5)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 7, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 6)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 8, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 7)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 9, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 8)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 10, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 9)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 11, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 10)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 12, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 11)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 13, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 12)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 14, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 13)
Command1.SetFocus

ElseIf Mid(buyamount.Text, 15, 1) = "." Then
buyamount.Text = Left(buyamount.Text, 14)
Command1.SetFocus

End If
End Sub

Private Sub Command1_Click()
On Error GoTo errhan

If buyamount.Text = "0" Or buyamount.Text = "" Then
GoTo singleunit
Else
GoTo multiunit

singleunit:

If selldeal.Text = ":cocaine" Then
If Val(pcoc.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
cocaine.Caption = cocaine.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pcoc.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":heroin" Then
If Val(pher.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
heroin.Caption = heroin.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pher.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":meth" Then
If Val(pmeth.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
meth.Caption = meth.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pmeth.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":dust" Then
If Val(pdust.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
dust.Caption = dust.Caption + 1
totcash.Caption = totcash.Caption - pdust.Caption
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":weed" Then
If Val(pweed.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
weed.Caption = weed.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pweed.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":ganja" Then
If Val(pgan.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
ganja.Caption = ganja.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pgan.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":smack" Then
If Val(psma.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
smack.Caption = smack.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(psma.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":crack" Then
If Val(pcra.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
crack.Caption = crack.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pcra.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":xtc" Then
If Val(pxtc.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
xtc.Caption = xtc.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pxtc.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":ice" Then
If Val(pice.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
ice.Caption = ice.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pice.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":speed" Then
If Val(pspe.Caption) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) = Val(maxroom.Caption) Or Val(room.Caption) > Val(maxroom.Caption) Then
roomfix
Else
speed.Caption = speed.Caption + 1
totcash.Caption = Val(totcash.Caption) - Val(pspe.Caption)
room.Caption = room.Caption + 1
buyamount.Text = ""
End If
End If
End If
Exit Sub

multiunit:

If selldeal.Text = ":cocaine" Then
If Val(pcoc.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pxtc.Caption * buyamount.Text)
cocaine.Caption = cocaine.Caption + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":heroin" Then
If Val(pher.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pher.Caption * buyamount.Text)
heroin.Caption = heroin.Caption + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":meth" Then
If Val(pmeth.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pmeth.Caption * buyamount.Text)
meth.Caption = meth.Caption + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":dust" Then
If Val(pdust.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pdust.Caption * buyamount.Text)
dust.Caption = dust.Caption + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":weed" Then
If Val(pweed.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pweed.Caption * buyamount.Text)
weed.Caption = weed.Caption + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":ganja" Then
If Val(pgan.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pgan.Caption * buyamount.Text)
ganja.Caption = ganja.Caption + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":smack" Then
If Val(psma.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(psma.Caption * buyamount.Text)
smack.Caption = Val(smack.Caption) + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":crack" Then
If Val(pcra.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pcra.Caption * buyamount.Text)
crack.Caption = Val(crack.Caption) + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":xtc" Then
If Val(pxtc.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pxtc.Caption * buyamount.Text)
xtc.Caption = Val(xtc.Caption) + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":ice" Then
If Val(pice.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pice.Caption * buyamount.Text)
ice.Caption = Val(ice.Caption) + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
buyamount.Text = ""
End If
End If

ElseIf selldeal.Text = ":speed" Then
If Val(pspe.Caption * buyamount.Text) > Val(totcash.Caption) Then
buyfix
Else
If Val(room.Caption) + Val(buyamount.Text) > Val(maxroom.Caption) Then
roomfix
Else
totcash.Caption = Val(totcash.Caption) - Val(pspe.Caption * buyamount.Text)
speed.Caption = Val(speed.Caption) + Val(buyamount.Text)
room.Caption = Val(room.Caption) + Val(buyamount.Text)
End If
End If
End If
End If
Exit Sub

errhan:
MsgBox "Invalid units to purchase.", vbExclamation, "Invalid"
buyamount.Text = ""
End Sub

Private Sub Command10_Click()
citybreak
Command10.Enabled = False
day
End Sub

Private Sub Command11_Click()
bankman.Show
End Sub

Private Sub Command13_Click()
sabank.Show
End Sub

Sub day()
buyamount.Text = ""
sellamount.Text = ""

If days.Caption = "1" Then
MsgBox "Chino Is enraged with you and has mutated. and wants to battle.", vbExclamation, "Battle"
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "PCHINO"
battle.Show
days.Caption = days.Caption - 1
Exit Sub
End If

If days.Caption = "20" Then
MsgBox "Chino Has come to collect his money, but also decided to battle you anyway.", vbExclamation, "Battle"
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "CHINO"
battle.Show
days.Caption = days.Caption - 1
Exit Sub
End If

If days.Caption = "50" Then
MsgBox "The Federal Police have released there new prototype Armoured Police officer, Armoured Cop has found you and wants to battle.", vbExclamation, "Battle"
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "ARMCOP"
battle.Show
days.Caption = days.Caption - 1
Exit Sub
End If

If days.Caption = "0" Then
MsgBox "Game over, score has been logged in high scores.", vbExclamation, "Game Over"
Me.Hide
Unload gunshop
Unload battle
Unload bankman
Unload tele
highscores.Show
Exit Sub
End If

If totdebt.Caption = "0" Then
totdebt.Caption = "0"
Else
totdebt.Caption = totdebt.Caption + 1000
End If
days.Caption = days.Caption - 1
randem
bttl

End Sub

Private Sub Command15_Click()
day
End Sub

Private Sub Command12_Click()
tele.Show
End Sub
Sub sellfix()
MsgBox "No Units to sell", vbExclamation, "No Units"
End Sub

Private Sub Command16_Click()
day
End Sub

Private Sub Command19_Click()

End Sub

Private Sub Command14_Click()
totcash.Caption = Val(totcash.Caption) - Val(Text1.Text)
End Sub

Private Sub Command2_Click()
On Error GoTo errhan


If sellamount.Text = "" Or sellamount.Text = "0" Then
GoTo singleunit
Else
GoTo multiunit
End If

singleunit:
If buydeal.Text = "%cocaine" Then
If cocaine.Caption = "0" Then
sellfix
Else
cocaine.Caption = cocaine.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pcoc.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%heroin" Then
If heroin.Caption = "0" Then
sellfix
Else
heroin.Caption = heroin.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pher.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%meth" Then
If meth.Caption = "0" Then
sellfix
Else
meth.Caption = meth.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pmeth.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%dust" Then
If dust.Caption = "0" Then
sellfix
Else
dust.Caption = dust.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pdust.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%weed" Then
If weed.Caption = "0" Then
sellfix
Else
weed.Caption = weed.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pweed.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%ganja" Then
If ganja.Caption = "0" Then
sellfix
Else
ganja.Caption = ganja.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pgan.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%smack" Then
If smack.Caption = "0" Then
sellfix
Else
smack.Caption = smack.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(psma.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%crack" Then
If crack.Caption = "0" Then
sellfix
Else
crack.Caption = crack.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pcra.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%xtc" Then
If xtc.Caption = "0" Then
sellfix
Else
xtc.Caption = xtc.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pxtc.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%ice" Then
If ice.Caption = "0" Then
sellfix
Else
ice.Caption = ice.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pice.Caption)
room.Caption = Val(room.Caption) - 1
End If

ElseIf buydeal.Text = "%speed" Then
If speed.Caption = "0" Then
sellfix
Else
speed.Caption = speed.Caption - 1
totcash.Caption = Val(totcash.Caption) + Val(pspe.Caption)
room.Caption = Val(room.Caption) - 1
End If
End If
Exit Sub

multiunit:
If buydeal.Text = "%cocaine" Then
 If Val(sellamount.Text) > Val(cocaine.Caption) Then
 sellfix
Else
cocaine.Caption = Val(cocaine.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pcoc.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%heroin" Then
 If Val(sellamount.Text) > Val(heroin.Caption) Then
 sellfix
Else
heroin.Caption = Val(heroin.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pher.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%meth" Then
 If Val(sellamount.Text) > Val(meth.Caption) Then
 sellfix
Else
meth.Caption = Val(meth.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pmeth.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%dust" Then
 If Val(sellamount.Text) > Val(dust.Caption) Then
 sellfix
Else
dust.Caption = Val(dust.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pdust.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%weed" Then
 If Val(sellamount.Text) > Val(weed.Caption) Then
 sellfix
Else
weed.Caption = Val(weed.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pweed.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%ganja" Then
 If Val(sellamount.Text) > Val(ganja.Caption) Then
 sellfix
Else
ganja.Caption = Val(ganja.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pgan.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%smack" Then
 If Val(sellamount.Text) > Val(smack.Caption) Then
 sellfix
Else
smack.Caption = Val(smack.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(psma.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%crack" Then
 If Val(sellamount.Text) > Val(crack.Caption) Then
 sellfix
Else
crack.Caption = Val(crack.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pcra.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%xtc" Then
 If Val(sellamount.Text) > Val(xtc.Caption) Then
 sellfix
Else
xtc.Caption = Val(xtc.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pxtc.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%ice" Then
 If Val(sellamount.Text) > Val(ice.Caption) Then
 sellfix
Else
ice.Caption = Val(ice.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pice.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
sellamount.Text = ""
End If

ElseIf buydeal.Text = "%speed" Then
 If Val(sellamount.Text) > Val(speed.Caption) Then
 sellfix
Else
speed.Caption = Val(speed.Caption) - Val(sellamount.Text)
totcash.Caption = Val(totcash.Caption) + Val(pspe.Caption * sellamount.Text)
room.Caption = Val(room.Caption) - sellamount.Text
End If
End If
Exit Sub

errhan:
MsgBox "Invalid Amount to sell.", vbExclamation, "Invalid"
sellamount.Text = ""
End Sub

Private Sub Command3_Click()
citybreak
Command3.Enabled = False
day
End Sub


Private Sub Command4_Click()
citybreak
Command4.Enabled = False
day
End Sub

Private Sub Command5_Click()
citybreak
Command5.Enabled = False
day
End Sub

Private Sub Command6_Click()
citybreak
Command6.Enabled = False
day
End Sub

Private Sub Command7_Click()
citybreak
Command7.Enabled = False
day
End Sub

Private Sub Command8_Click()
citybreak
Command8.Enabled = False
day
End Sub

Private Sub Command9_Click()
citybreak
Command9.Enabled = False
day
End Sub

Private Sub days_Change()
If Val(days.Caption) > Val(maxdays.Caption) Then
days.Caption = maxdays.Caption
End If
End Sub
Sub citybreak()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
End Sub
Private Sub Form_Load()
Call Label3_Click
Call Label45_Click
Me.Show
buy_hider
sell_hider
randem
a1.Visible = True
b1.Visible = True
End Sub
Sub sell_hider()
a1.Visible = False
a2.Visible = False
a3.Visible = False
a4.Visible = False
a5.Visible = False
a6.Visible = False
a7.Visible = False
a8.Visible = False
a9.Visible = False
a10.Visible = False
a11.Visible = False

End Sub

Private Sub Form_Terminate()
Unload battle
Unload bankman
Unload cheats
Unload gunshop
Unload sabank
Unload tele
Unload splash
Unload about
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload battle
Unload bankman
Unload cheats
Unload gunshop
Unload sabank
Unload tele
Unload splash
Unload about
End Sub

Private Sub Label14_Click()
buy ("%heroin")
buy_hider
b2.Visible = True
End Sub

Private Sub Label15_Click()
buy ("%meth")
buy_hider
b3.Visible = True
End Sub

Private Sub Label16_Click()
buy ("%ice")
buy_hider
b10.Visible = True
End Sub

Private Sub Label17_Click()
buy ("%weed")
buy_hider
b5.Visible = True
End Sub

Private Sub Label18_Click()
buy ("%ganja")
buy_hider
b6.Visible = True
End Sub

Private Sub Label19_Click()
buy ("%dust")
buy_hider
b4.Visible = True
End Sub

Private Sub Label20_Click()
buy ("%crack")
buy_hider
b8.Visible = True
End Sub

Private Sub Label21_Click()
buy ("%smack")
buy_hider
b7.Visible = True
End Sub

Private Sub Label22_Click()
buy ("%xtc")
buy_hider
b9.Visible = True
End Sub

Private Sub Label23_Click()
buy ("%speed")
buy_hider
b11.Visible = True
End Sub

Private Sub Label26_Click()
If buyamount.Text = "CHEAT" And sellamount.Text = "MENU" Then
cheats.Show
End If
End Sub

Private Sub Label3_Click()
buy ("%cocaine")
buy_hider
b1.Visible = True
End Sub
Sub buy(drug)
buydeal.Text = drug
End Sub
Sub buy_hider()
b1.Visible = False
b2.Visible = False
b3.Visible = False
b4.Visible = False
b5.Visible = False
b6.Visible = False
b7.Visible = False
b8.Visible = False
b9.Visible = False
b10.Visible = False
b11.Visible = False
End Sub

Private Sub Label35_Click()
sell (":speed")
sell_hider
a11.Visible = True
End Sub

Private Sub Label36_Click()
sell (":xtc")
sell_hider
a9.Visible = True
End Sub

Private Sub Label37_Click()
sell (":smack")
sell_hider
a7.Visible = True
End Sub

Private Sub Label38_Click()
sell (":crack")
sell_hider
a8.Visible = True
End Sub

Private Sub Label39_Click()
sell (":dust")
sell_hider
a4.Visible = True
End Sub

Private Sub Label40_Click()
sell (":ganja")
sell_hider
a6.Visible = True
End Sub

Private Sub Label41_Click()
sell (":weed")
sell_hider
a5.Visible = True
End Sub

Private Sub Label42_Click()
sell (":ice")
sell_hider
a10.Visible = True
End Sub

Private Sub Label43_Click()
sell (":meth")
sell_hider
a3.Visible = True
End Sub

Private Sub Label44_Click()
sell (":heroin")
sell_hider
a2.Visible = True
End Sub

Private Sub Label45_Click()
sell (":cocaine")
sell_hider
a1.Visible = True
End Sub
Sub sell(drug)
selldeal.Text = drug
End Sub



Private Sub Label5_Click()
about.Show
End Sub

Private Sub medi_Click()
If medikits.Text = "0" Then
MsgBox "You have no Medi-Kits in Inventory", vbExclamation, "Medi-Kits"
Else
medikits.Text = medikits.Text - 1
medi.Caption = "Medi-Kit " & medikits.Text
health.Caption = mhealth.Caption
End If
End Sub

Private Sub price_Click()
streetdope.SetFocus
End Sub

Private Sub price_GotFocus()
streetdope.SetFocus
End Sub




Private Sub mhealth_Change()
If Val(mhealth.Caption) > "99999999" Then
mhealth.Caption = "99999999"
End If
End Sub

Private Sub sellamount_Change()
If Left(sellamount.Text, 1) = "-" Then
sellamount.Text = Right(sellamount.Text, Len(sellamount.Text) - 1)
End If

If Mid(sellamount.Text, 2, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 1)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 3, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 2)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 4, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 3)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 5, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 4)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 6, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 5)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 7, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 6)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 8, 1) = "." Then
sellamount = Left(sellamount.Text, 7)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 9, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 8)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 10, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 9)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 11, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 10)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 12, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 11)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 13, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 12)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 14, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 13)
Command2.SetFocus

ElseIf Mid(sellamount.Text, 15, 1) = "." Then
sellamount.Text = Left(sellamount.Text, 14)
Command2.SetFocus

End If
End Sub

Private Sub selldeal_GotFocus()
Command2.SetFocus
End Sub

Private Sub selldeal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
selldeal.MousePointer = vbArrow
End Sub


Private Sub ster_Click()
If steroid.Text = "0" Then
ster.Caption = "Steroids 0"
MsgBox "You have no Steroids in Inventory", vbExclamation, "Steroids"
Else
steroid.Text = steroid.Text - 1
ster.Caption = "Steroids " & steroid.Text
mhealth.Caption = mhealth + 20
End If
End Sub



Private Sub totcash_Change()
If Val(totcash.Caption) > Val(secret.Text) Then
totcash.Caption = Val(secret.Text)
End If
If Val(totcash.Caption) < 0 Then
totcash.Caption = "0"
End If

If Mid(totcash.Caption, 2, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 1)

ElseIf Mid(totcash.Caption, 3, 1) = "." Then
totcash.Caption = Left(totcash, 2)

ElseIf Mid(totcash.Caption, 4, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 3)

ElseIf Mid(totcash.Caption, 5, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 4)

ElseIf Mid(totcash.Caption, 6, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 5)

ElseIf Mid(totcash.Caption, 7, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 6)

ElseIf Mid(totcash.Caption, 8, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 7)

ElseIf Mid(totcash.Caption, 9, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 8)

ElseIf Mid(totcash.Caption, 10, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 9)

ElseIf Mid(totcash.Caption, 11, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 10)

ElseIf Mid(totcash.Caption, 12, 1) = "." Then
totcash.Caption = Left(totcash.Caption, 11)

End If
End Sub


Private Sub weap_Click()
gunshop.Show
End Sub


Sub randem()
' random price maker
Randomize

X = Int(Rnd * 30)
Select Case X
Case 0
pweed.Caption = "85"
Case 1
pweed.Caption = "180"
Case 2
pweed.Caption = "378"
Case 3
pweed.Caption = "150"
Case 4
pweed.Caption = "200"
Case 5
pweed.Caption = "890"
MsgBox "Whoa! Cops burnt up a couple of plantations. Weed prices boom!", vbExclamation, "Weed"
Case 6
pweed.Caption = "425"
Case 7
pweed.Caption = "235"
Case 8
pweed.Caption = "546"
Case 9
pweed.Caption = "23"
MsgBox "Weed Prices dropped, Chino lost a whole crop to Rival Drug Lords", vbExclamation, "Weed"
Case 10
pweed.Caption = "39"
MsgBox "Weed Prices dropped, Chino lost a whole crop to Rival Drug Lords", vbExclamation, "Weed"
Case 11
pweed.Caption = "298"
Case 12
pweed.Caption = "462"
Case 13
pweed.Caption = "221"
Case 14
pweed.Caption = "548"
Case 15
pweed.Caption = "235"
Case 16
pweed.Caption = "834"
Case 17
pweed.Caption = "346"
Case 18
pweed.Caption = "924"
Case 19
pweed.Caption = "478"
Case 20
pweed.Caption = "389"
Case 21
pweed.Caption = "286"
Case 22
pweed.Caption = "325"
Case 23
pweed.Caption = "532"
Case 24
pweed.Caption = "789"
Case 25
pweed.Caption = "280"
Case 26
pweed.Caption = "190"
Case 27
pweed.Caption = "260"
Case 28
pweed.Caption = "545"
Case 29
pweed.Caption = "213"
End Select

X = Int(Rnd * 30)
Select Case X
Case 0
pcoc.Caption = "180578"
MsgBox "good time to sell Cocaine.", vbExclamation, "Cocaine"
Case 1
pcoc.Caption = "263200"
MsgBox "you are the only man in town with Cocaine. buyers are willing to pay crazy prices.", vbExclamation, "Cocaine"
Case 2
pcoc.Caption = "27920"
Case 3
pcoc.Caption = "23560"
Case 4
pcoc.Caption = "34785"
Case 5
pcoc.Caption = "60832"
Case 6
pcoc.Caption = "42912"
Case 7
pcoc.Caption = "31817"
Case 8
pcoc.Caption = "16459"
Case 9
pcoc.Caption = "13891"
MsgBox "Thieves sell cheap Cocaine stolen from Chino's Labs.", vbExclamation, "Cocaine"
Case 10
pcoc.Caption = "65952"
Case 11
pcoc.Caption = "39485"
Case 12
pcoc.Caption = "11430"
MsgBox "Thieves sell cheap Cocaine stolen from Chino's Labs.", vbExclamation, "Cocaine"
Case 13
pcoc.Caption = "82353"
Case 14
pcoc.Caption = "72290"
Case 15
pcoc.Caption = "58342"
Case 16
pcoc.Caption = "43685"
Case 17
pcoc.Caption = "64458"
Case 18
pcoc.Caption = "89432"
Case 19
pcoc.Caption = "35763"
Case 20
pcoc.Caption = "34265"
Case 21
pcoc.Caption = "53562"
Case 22
pcoc.Caption = "32573"
Case 23
pcoc.Caption = "53281"
Case 24
pcoc.Caption = "78926"
Case 25
pcoc.Caption = "28032"
Case 26
pcoc.Caption = "19096"
Case 27
pcoc.Caption = "26033"
Case 28
pcoc.Caption = "54564"
Case 29
pcoc.Caption = "21300"
End Select

X = Int(Rnd * 20)
Select Case X
Case 0
pher.Caption = "10592"
Case 1
pher.Caption = "15672"
Case 2
pher.Caption = "454062"
MsgBox "Armor Cop showed up to the warehouse, Chino isn't going to be happy of that", vbExclamation, "Heroin"
Case 3
pher.Caption = "12056"
Case 4
pher.Caption = "18577"
Case 5
pher.Caption = "12434"
Case 6
pher.Caption = "16354"
Case 7
pher.Caption = "8894"
MsgBox "Heroin is in good supply, buy some now to make quick cash", vbExclamation, "Heroin"
Case 8
pher.Caption = "13433"
Case 9
pher.Caption = "14675"
Case 10
pher.Caption = "17034"
Case 11
pher.Caption = "17782"
Case 12
pher.Caption = "14221"
Case 13
pher.Caption = "15390"
Case 14
pher.Caption = "5605"
MsgBox "Chino is sending rival dealers down the drain, selling stolen Heroin", vbExclamation, "Heroin"
Case 15
pher.Caption = "34265"
Case 16
pher.Caption = "84245"
Case 17
pher.Caption = "56722"
Case 18
pher.Caption = "25474"
Case 19
pher.Caption = "95467"
End Select

X = Int(Rnd * 30)
Select Case X
Case 0
pmeth.Caption = "2367"
Case 1
pmeth.Caption = "3461"
Case 2
pmeth.Caption = "3546"
Case 3
pmeth.Caption = "5232"
Case 4
pmeth.Caption = "560"
MsgBox "Rival dealers are trying to get many sales today!", vbExclamation, "Crystal Meth"
Case 5
pmeth.Caption = "3312"
Case 6
pmeth.Caption = "3634"
Case 7
pmeth.Caption = "4612"
Case 8
pmeth.Caption = "2894"
Case 9
pmeth.Caption = "2760"
Case 10
pmeth.Caption = "4256"
Case 11
pmeth.Caption = "899"
MsgBox "Chino must have found the home recipe for Crystal Meth, prices are cheap!", vbExclamation, "Crystal Meth"
Case 12
pmeth.Caption = "2859"
Case 13
pmeth.Caption = "3462"
Case 14
pmeth.Caption = "2234"
Case 15
pmeth.Caption = "1265"
Case 16
pmeth.Caption = "2723"
Case 17
pmeth.Caption = "1375"
Case 18
pmeth.Caption = "2167"
Case 19
pmeth.Caption = "7432"
Case 20
pmeth.Caption = "3894"
Case 21
pmeth.Caption = "3863"
Case 22
pmeth.Caption = "7252"
Case 23
pmeth.Caption = "4321"
Case 24
pmeth.Caption = "5897"
Case 25
pmeth.Caption = "2808"
Case 26
pmeth.Caption = "5233"
Case 27
pmeth.Caption = "2344"
Case 28
pmeth.Caption = "5451"
Case 29
pmeth.Caption = "2134"
End Select


X = Int(Rnd * 30)
Select Case X
Case 0
pdust.Caption = "5665"
Case 1
pdust.Caption = "8904"
Case 2
pdust.Caption = "6783"
Case 3
pdust.Caption = "7480"
Case 4
pdust.Caption = "4565"
Case 5
pdust.Caption = "1982"
MsgBox "With Angel Dust so cheap, I reckon a few Angels fell down to earth.", vbExclamation, "Angel Dust"
Case 6
pdust.Caption = "58933"
MsgBox "Whoa! Major Dust Bust!", vbExclamation, "Angel Dust"
Case 7
pdust.Caption = "4865"
Case 8
pdust.Caption = "8675"
Case 9
pdust.Caption = "7568"
Case 10
pdust.Caption = "9734"
Case 11
pdust.Caption = "5289"
Case 12
pdust.Caption = "5823"
Case 13
pdust.Caption = "7553"
Case 14
pdust.Caption = "8902"
Case 15
pdust.Caption = "12568"
Case 16
pdust.Caption = "35743"
Case 17
pdust.Caption = "1235"
Case 18
pdust.Caption = "8412"
Case 19
pdust.Caption = "6546"
Case 20
pdust.Caption = "4389"
Case 21
pdust.Caption = "8286"
Case 22
pdust.Caption = "3825"
Case 23
pdust.Caption = "4532"
Case 24
pdust.Caption = "7849"
Case 25
pdust.Caption = "2480"
Case 26
pdust.Caption = "7190"
Case 27
pdust.Caption = "8450"
Case 28
pdust.Caption = "5445"
Case 29
pdust.Caption = "6312"
End Select

X = Int(Rnd * 30)
Select Case X
Case 0
pgan.Caption = "85"
Case 1
pgan.Caption = "180"
Case 2
pgan.Caption = "378"
Case 3
pgan.Caption = "150"
Case 4
pgan.Caption = "200"
Case 5
pgan.Caption = "890"
MsgBox "Fresh from Jamaica, super Ganja going at high prices.", vbExclamation, "Ganga"
Case 6
pgan.Caption = "425"
Case 7
pgan.Caption = "235"
Case 8
pgan.Caption = "546"
Case 9
pgan.Caption = "456"
Case 10
pgan.Caption = "39"
MsgBox "Ganga Prices dropped, Chino lost a whole crop to Rival Drug Lords", vbExclamation, "Ganga"
Case 11
pgan.Caption = "298"
Case 12
pgan.Caption = "462"
Case 13
pgan.Caption = "221"
Case 14
pgan.Caption = "548"
Case 15
pgan.Caption = "324"
Case 16
pgan.Caption = "357"
Case 17
pgan.Caption = "784"
Case 18
pgan.Caption = "264"
Case 19
pgan.Caption = "454"
Case 20
pgan.Caption = "579"
Case 21
pgan.Caption = "486"
Case 22
pgan.Caption = "13595"
Case 23
pgan.Caption = "532"
Case 24
pgan.Caption = "1089"
Case 25
pgan.Caption = "1146"
Case 26
pgan.Caption = "1001"
Case 27
pgan.Caption = "543"
Case 28
pgan.Caption = "545"
Case 29
pgan.Caption = "213"
End Select

X = Int(Rnd * 30)
Select Case X
Case 0
psma.Caption = "32567"
Case 1
psma.Caption = "28579"
Case 2
psma.Caption = "25344"
Case 3
psma.Caption = "35649"
Case 4
psma.Caption = "26994"
Case 5
psma.Caption = "8933"
MsgBox "Imported smack selling cheap!", vbExclamation, "Smack"
Case 6
psma.Caption = "34561"
Case 7
psma.Caption = "29086"
Case 8
psma.Caption = "31389"
Case 9
psma.Caption = "32561"
Case 10
psma.Caption = "165982"
MsgBox "Drug bust, by Armor Cop. Prices BOOM!", vbExclamation, "Smack"
Case 11
psma.Caption = "27845"
Case 12
psma.Caption = "30730"
Case 13
psma.Caption = "27506"
Case 14
psma.Caption = "33098"
Case 15
psma.Caption = "56723"
Case 16
psma.Caption = "58322"
Case 17
psma.Caption = "36554"
Case 18
psma.Caption = "45648"
Case 19
psma.Caption = "56549"
Case 20
psma.Caption = "38789"
Case 21
psma.Caption = "28658"
Case 22
psma.Caption = "32524"
Case 23
psma.Caption = "56532"
Case 24
psma.Caption = "78219"
Case 25
psma.Caption = "32880"
Case 26
psma.Caption = "61520"
Case 27
psma.Caption = "22354"
Case 28
psma.Caption = "65424"
Case 29
psma.Caption = "46134"
End Select

X = Int(Rnd * 30)
Select Case X
Case 0
pcra.Caption = "27567"
Case 1
pcra.Caption = "23579"
Case 2
pcra.Caption = "20344"
Case 3
pcra.Caption = "30649"
Case 4
pcra.Caption = "21994"
Case 5
pcra.Caption = "7933"
MsgBox "Pakistan crack selling cheap!", vbExclamation, "Crack"
Case 6
pcra.Caption = "29561"
Case 7
pcra.Caption = "24086"
Case 8
pcra.Caption = "26389"
Case 9
pcra.Caption = "27561"
Case 10
pcra.Caption = "60982"
MsgBox "Malasian boat carrying crack sinks! Prices soar!", vbExclamation, "Crack"
Case 11
pcra.Caption = "22845"
Case 12
pcra.Caption = "25730"
Case 13
pcra.Caption = "22506"
Case 14
pcra.Caption = "28098"
Case 15
pcra.Caption = "32342"
Case 16
pcra.Caption = "16541"
Case 17
pcra.Caption = "45414"
Case 18
pcra.Caption = "46463"
Case 19
pcra.Caption = "19144"
Case 20
pcra.Caption = "38119"
Case 21
pcra.Caption = "28642"
Case 22
pcra.Caption = "32215"
Case 23
pcra.Caption = "53422"
Case 24
pcra.Caption = "78429"
Case 25
pcra.Caption = "28445"
Case 26
pcra.Caption = "23560"
Case 27
pcra.Caption = "29146"
Case 28
pcra.Caption = "54535"
Case 29
pcra.Caption = "27374"
End Select

X = Int(Rnd * 20)
Select Case X
Case 0
pxtc.Caption = "28"
Case 1
pxtc.Caption = "23"
Case 2
pxtc.Caption = "15"
Case 3
pxtc.Caption = "42"
Case 4
pxtc.Caption = "890"
MsgBox "XTC bust in all nightclubs across the world", vbExclamation, "XTC"
Case 5
pxtc.Caption = "36"
Case 6
pxtc.Caption = "25"
Case 7
pxtc.Caption = "35"
Case 8
pxtc.Caption = "46"
Case 9
pxtc.Caption = "1"
MsgBox "Rival Drug Lords *find* XTC", vbExclamation, "XTC"
Case 10
pxtc.Caption = "39"
Case 11
pxtc.Caption = "18"
Case 12
pxtc.Caption = "42"
Case 13
pxtc.Caption = "21"
Case 14
pxtc.Caption = "38"
Case 15
pxtc.Caption = "19"
Case 16
pxtc.Caption = "86"
Case 17
pxtc.Caption = "72"
Case 18
pxtc.Caption = "68"
Case 19
pxtc.Caption = "38"
End Select


X = Int(Rnd * 30)
Select Case X
Case 0
pice.Caption = "28996"
Case 1
pice.Caption = "35890"
Case 2
pice.Caption = "38932"
Case 3
pice.Caption = "32032"
Case 4
pice.Caption = "75062"
MsgBox "IcE melts as prices rise!", vbExclamation, "IcE"
Case 5
pice.Caption = "36345"
Case 6
pice.Caption = "25421"
Case 7
pice.Caption = "35073"
Case 8
pice.Caption = "31864"
Case 9
pice.Caption = "15092"
MsgBox "Cheap Home made IcE pulled out form Chino's Ring.", vbExclamation, "IcE"
Case 10
pice.Caption = "39345"
Case 11
pice.Caption = "23465"
Case 12
pice.Caption = "28934"
Case 13
pice.Caption = "28990"
Case 14
pice.Caption = "33782"
Case 15
pice.Caption = "54465"
Case 16
pice.Caption = "42864"
Case 17
pice.Caption = "36542"
Case 18
pice.Caption = "46545"
Case 19
pice.Caption = "23434"
Case 20
pice.Caption = "51564"
Case 21
pice.Caption = "13300"
Case 22
pice.Caption = "45310"
Case 23
pice.Caption = "37561"
Case 24
pice.Caption = "45268"
Case 25
pice.Caption = "64285"
Case 26
pice.Caption = "58497"
Case 27
pice.Caption = "52631"
Case 28
pice.Caption = "54553"
Case 29
pice.Caption = "35213"
End Select

X = Int(Rnd * 30)
Select Case X
Case 0
pspe.Caption = "566"
Case 1
pspe.Caption = "890"
Case 2
pspe.Caption = "678"
Case 3
pspe.Caption = "748"
Case 4
pspe.Caption = "456"
Case 5
pspe.Caption = "198"
MsgBox "Speed going Fast!", vbExclamation, "Speed"
Case 6
pspe.Caption = "5893"
MsgBox "Whoa! New speed bumps in place!", vbExclamation, "Speed"
Case 7
pspe.Caption = "486"
Case 8
pspe.Caption = "867"
Case 9
pspe.Caption = "756"
Case 10
pspe.Caption = "973"
Case 11
pspe.Caption = "528"
Case 12
pspe.Caption = "582"
Case 13
pspe.Caption = "755"
Case 14
pspe.Caption = "890"
Case 15
pspe.Caption = "967"
Case 16
pspe.Caption = "813"
Case 17
pspe.Caption = "664"
Case 18
pspe.Caption = "654"
Case 19
pspe.Caption = "547"
Case 20
pspe.Caption = "389"
Case 21
pspe.Caption = "386"
Case 22
pspe.Caption = "425"
Case 23
pspe.Caption = "532"
Case 24
pspe.Caption = "789"
Case 25
pspe.Caption = "280"
Case 26
pspe.Caption = "790"
Case 27
pspe.Caption = "660"
Case 28
pspe.Caption = "563"
Case 29
pspe.Caption = "730"
End Select

End Sub

Sub bttl()

Randomize

X = Int(Rnd * 50)
Select Case X
Case 1
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "JOE"
battle.Show
Case 2

Case 3

Case 4

Case 5
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "TEX"
battle.Show
Case 6

Case 7

Case 8
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "COP"
battle.Show
Case 9

Case 10

Case 11
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "JOE"
battle.Show
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
Case 12

Case 13

Case 14

Case 15

Case 16

Case 17
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "TEX"
battle.Show
Case 18

Case 19

Case 20

Case 21

Case 22

Case 23

Case 24
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "JOE"
battle.Show
Case 25

Case 26

Case 27

Case 28
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "COP"
battle.Show
Case 29

Case 30

Case 31

Case 32

Case 33
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "TEX"
battle.Show
Case 34

Case 35

Case 36

Case 37
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "JOE"
battle.Show
Case 38

Case 39

Case 40

Case 41

Case 42

Case 43
tele.Enabled = False
gunshop.Enabled = False
bankman.Enabled = False
sabank.Enabled = False
inven.Enabled = False
locations.Enabled = False
battle.attacker.Text = "JOE"
battle.Show
Case 44

Case 45

Case 46

Case 47

Case 48

Case 49

End Select
End Sub
