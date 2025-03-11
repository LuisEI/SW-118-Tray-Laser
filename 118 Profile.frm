VERSION 5.00
Begin VB.Form frmProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "106 ATC Alase Technologies - COM Interface Tester Text Profile"
   ClientHeight    =   11445
   ClientLeft      =   -165
   ClientTop       =   105
   ClientWidth     =   14790
   Icon            =   "118 Profile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11445
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option6 
      Caption         =   "[5] ATC"
      Height          =   360
      Left            =   8145
      TabIndex        =   108
      Top             =   840
      Width           =   1000
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[4] Logo"
      Height          =   360
      Left            =   7140
      TabIndex        =   107
      Top             =   840
      Width           =   1000
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[3] Text 4"
      Height          =   360
      Left            =   6135
      TabIndex        =   106
      Top             =   840
      Width           =   1000
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[2] Text 3"
      Height          =   360
      Left            =   5130
      TabIndex        =   105
      Top             =   840
      Width           =   1000
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[1] Text 2"
      Height          =   360
      Left            =   4125
      TabIndex        =   104
      Top             =   840
      Width           =   1000
   End
   Begin VB.OptionButton option1 
      Caption         =   "[0] Text 1"
      Height          =   360
      Left            =   3120
      TabIndex        =   103
      Top             =   840
      Value           =   -1  'True
      Width           =   1000
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   255
      Left            =   11400
      TabIndex        =   100
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Copy Profile"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   " Scale "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8520
      TabIndex        =   91
      Top             =   7680
      Width           =   3135
      Begin VB.CommandButton cmdScale 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ScaleObj"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   93
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   92
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label46 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   98
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label45 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   97
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label44 
         Caption         =   "YScale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   96
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label43 
         Caption         =   "XScale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   95
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSetAll 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SetAllObjProfile()"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMM3 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   89
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdPP3 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      TabIndex        =   88
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdP3 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   87
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton cmdM3 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6735
      TabIndex        =   86
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton cmdMM2 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   85
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdPP2 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      TabIndex        =   84
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdP2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   83
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdM2 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6735
      TabIndex        =   82
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdM1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6735
      TabIndex        =   81
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdP1 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   80
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdPP1 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      TabIndex        =   79
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdMM1 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   78
      Top             =   2520
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [FIXTURE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\106 MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FIXTURE"
      Top             =   1320
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.TextBox txtInch 
      Height          =   300
      Left            =   7800
      TabIndex        =   73
      Text            =   "0"
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Inches to Bits"
      Height          =   300
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   " Rotate "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   11760
      TabIndex        =   67
      Top             =   7680
      Width           =   2655
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "[1] RotateObj()"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   68
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "degrees"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   70
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Angle :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   69
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.TextBox txtOrientation 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11280
      TabIndex        =   65
      Text            =   "0"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton btnMarkJob3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "[3] Mark Object"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   240
      Width           =   1875
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Exit to Main"
      Height          =   375
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   10560
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   " Position "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   480
      TabIndex        =   52
      Top             =   7680
      Width           =   4215
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   55
         Text            =   "100"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   54
         Text            =   "50"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "[5] SetObjPos()"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "(+-32768,+-32768)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   63
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Position <> (in)  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   60
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "Position ^ (in)  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   59
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "[HPosition]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   58
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "[VPosition]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   57
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "[0,0] Center Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   56
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Size "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4800
      TabIndex        =   44
      Top             =   7680
      Width           =   3615
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   47
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   46
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmd6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "[6]SetObjSize()"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label30 
         Caption         =   "Size <> (in)  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   51
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Size ^ (in)  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   50
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "[HSize]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   49
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "[VSize]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   48
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.TextBox txtObjIndex 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   42
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "[7] GetMulMatrixString"
      Height          =   300
      Left            =   12480
      TabIndex        =   41
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "[4] SetObjCharString()"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "[3] GetObjCharString()"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8880
      TabIndex        =   38
      Text            =   "1234A"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "[2] SetObjProfile()"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "[1] GetObjProfile()"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Wobblesize"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9120
      TabIndex        =   22
      Text            =   "X"
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Wobblefrequency"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9120
      TabIndex        =   20
      Text            =   "X"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Polygondelay"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   18
      Text            =   "XX"
      ToolTipText     =   "Polygondelay"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Markdelay"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   16
      Text            =   "XXX"
      ToolTipText     =   "Markdelay"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Jumpdelay"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9120
      TabIndex        =   13
      Text            =   "XXX"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Jumpspeed"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9120
      TabIndex        =   12
      Text            =   "XX"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Laseroffdelay"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   9
      Text            =   "XXX"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      DataField       =   "Laserondelay"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   8
      Text            =   "XXX"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtPulseWidth 
      BackColor       =   &H00FFC0FF&
      DataField       =   "PulseWidth"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   5
      Text            =   "X"
      ToolTipText     =   "PulseWidth"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtMarkspeed 
      BackColor       =   &H00FFC0FF&
      DataField       =   "Markspeed"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   4
      Text            =   "XXXXXX"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      DataField       =   "LaserPower"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   3
      Text            =   "XXXXXX"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtFrequency 
      BackColor       =   &H00FFC0FF&
      DataField       =   "Frequency"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   2
      Text            =   "XXX"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FROM [FIXTURE]"
      Height          =   300
      Left            =   8640
      TabIndex        =   109
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblMatrix_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M_ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7200
      TabIndex        =   102
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label48 
      Caption         =   "Matrix_ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   101
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label42 
      Caption         =   "Laser GST Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   77
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label41 
      Caption         =   "Laser Power Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   76
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label40 
      Caption         =   "[1 inch = 5842 bits]"
      Height          =   300
      Left            =   9960
      TabIndex        =   75
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label lblBit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   8880
      TabIndex        =   74
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label29 
      Caption         =   "Mark Orientation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9720
      TabIndex        =   66
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Valid Range :[-2,147,483,647 to 2,147,483,647]"
      Height          =   300
      Left            =   960
      TabIndex        =   62
      Top             =   7080
      Width           =   4455
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   12480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   12480
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label37 
      Caption         =   "Object Index :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   43
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label26 
      Caption         =   "[2 to 65535]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   35
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label25 
      Caption         =   "[0.02 to 250.0]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   34
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "[2 to 8000]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   33
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label23 
      Caption         =   "[-8000 to 8000]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   32
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label22 
      Caption         =   "[0 to 65500]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   31
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label21 
      Caption         =   "[0 to 65500]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   30
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "[0 to 100]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   29
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "[0 to 65500]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   28
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "[50 to 30000]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   27
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "[0 to 30000]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   26
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "[0 to 5000]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   25
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "[0 to 6000]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Wobble Width (bits)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   23
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Wobble Frequency (Hz)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   21
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "Poly Delay (us)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   19
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Mark Delay  (us)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   17
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Jump Delay  (us)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   15
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Jump Speed (bits/ms)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   14
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Laser Off Delay (us)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   11
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Laser On Delay (us)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   10
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Pulse Width (us)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Mark Speed (in/ms)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Laser Power (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Frequency (kHz)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnMarkJob3_Click()

' 3 IF OBJECT COUNT IS NOT ZERO

ObjIndex = CLng(txtObjIndex.Text)

BusyFlag = 1
While BusyFlag = 1
    AutomationInterface.GetBusyStatus 0, BusyFlag
Wend
AutomationInterface.MarkObj ObjIndex, Val(txtOrientation.Text)
 

MsgBox "Complete ObjIndex = " & ObjIndex, vbInformation, "Laser"

End Sub

Private Sub cmd1_Click()

'[1] GetObjProfile()

ObjIndex = CLng(txtObjIndex.Text)

ProfileIndex = 0

'ProfileIndex
'Markspeed
'Jumpspeed
'Jumpdelay
'Markdelay
'Polygondelay
'Laserpower
'Laseroffdelay
'Laserondelay
'Eightbitword, T1, T2
'ZAxis
'Varijumpdelay
'Varijumplength
'Wobblesize
'Wobblefrequency
'Powerreset
'Varipolydelay

AutomationInterface.GetObjProfile ObjIndex, ProfileIndex, Markspeed, Jumpspeed, Jumpdelay, Markdelay, Polygondelay, Laserpower, Laseroffdelay, Laserondelay, Eightbitword, T1, T2, zAxis, Varijumpdelay, Varijumplength, Wobblesize, Wobblefrequency, Powerreset, Varipolydelay

txtMarkspeed.Text = Markspeed
Text7.Text = Jumpspeed
Text8.Text = Jumpdelay
Text10.Text = Polygondelay
Text9.Text = Markdelay
Text4.Text = Laserpower
Text3.Text = Laseroffdelay
Text6.Text = Laserondelay
Text11.Text = Wobblesize
Text12.Text = Wobblefrequency
txtFrequency.Text = T1
txtPulseWidth.Text = T2

End Sub

Private Sub cmd2_Click()

'[2] SetObjProfile()

ObjIndex = CLng(txtObjIndex.Text)

ProfileIndex = 0

Markspeed = txtMarkspeed.Text
Jumpspeed = Text7.Text

Jumpdelay = Text8.Text
Polygondelay = Text10.Text
Markdelay = Text9.Text
Laserpower = Text4.Text
Laseroffdelay = Text3.Text
Laserondelay = Text6.Text
Wobblesize = Text11.Text
Wobblefrequency = Text12.Text
T1 = txtFrequency.Text
T2 = txtPulseWidth.Text

AutomationInterface.SetObjProfile ObjIndex, ProfileIndex, Markspeed, Jumpspeed, Jumpdelay, Markdelay, Polygondelay, Laserpower, Laseroffdelay, Laserondelay, Eightbitword, T1, T2, zAxis, Varijumpdelay, Varijumplength, Wobblesize, Wobblefrequency, Powerreset, Varipolydelay

End Sub

Private Sub cmd3_Click()

'[3] GetObjCharString()

ObjIndex = CLng(txtObjIndex.Text)

Dim sBuff As String

AutomationInterface.GetObjCharString ObjIndex, sBuff

Text13.Text = sBuff

End Sub

Private Sub cmd4_Click()

ObjIndex = CLng(txtObjIndex.Text)

Dim sBuff As String
sBuff = Text13.Text

AutomationInterface.SetObjCharString ObjIndex, sBuff

End Sub

Private Sub cmd5_Click()

'[5] SetObjPos()

ObjIndex = CLng(txtObjIndex.Text)

Dim HPosition As Long
Dim VPosition As Long

HPosition = Format(Val(Text14.Text) * INCHES_TO_BITS, "0")
VPosition = Format(Val(Text15.Text) * INCHES_TO_BITS, "0")

AutomationInterface.SetObjPos ObjIndex, HPosition, VPosition

End Sub

Private Sub cmd6_Click()

'[6] SetObSize()

ObjIndex = CLng(txtObjIndex.Text)

Dim HSize As Long
Dim VSize As Long

HSize = Format(Val(Text16.Text) * INCHES_TO_BITS, "0")
VSize = Format(Val(Text17.Text) * INCHES_TO_BITS, "0")

AutomationInterface.SetObjSize ObjIndex, HSize, VSize

End Sub

Private Sub cmd7_Click()

Dim a As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim X As Single
Dim Y As Single

ObjIndex = CLng(txtObjIndex.Text)

AutomationInterface.GetMulMatrixString ObjIndex, a, B, C, D, X, Y

Debug.Print a
Debug.Print B
Debug.Print C
Debug.Print D
Debug.Print X
Debug.Print Y

Beep
End Sub

Private Sub cmdConvert_Click()

lblBit.Caption = Val(txtInch.Text) * 5842

End Sub

Private Sub cmdCopy_Click()

Dim sSQL As String
Set FR_Database = OpenDatabase(ATC_LASER_BD)
sSQL = "SELECT * FROM [FIXTURE] WHERE [MATRIX ID] =" & MATRIX_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

ProfileIndex = 0

Markspeed = FR_Table.Fields("[Markspeed]")
Jumpspeed = FR_Table.Fields("[Jumpspeed]")
Jumpdelay = FR_Table.Fields("[Jumpdelay]")
Polygondelay = FR_Table.Fields("[Polygondelay]")
Markdelay = FR_Table.Fields("[Markdelay]")
Laserpower = FR_Table.Fields("[Laserpower]")
Laseroffdelay = FR_Table.Fields("[Laseroffdelay]")
Laserondelay = FR_Table.Fields("[Laserondelay]")
Wobblesize = FR_Table.Fields("[Wobblesize]")
Wobblefrequency = FR_Table.Fields("[Wobblefrequency]")
T1 = FR_Table.Fields("[Frequency]")
T2 = FR_Table.Fields("[PulseWidth]")

sSQL = "SELECT * FROM [FIXTURE] "
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Do Until FR_Table.EOF
        FR_Table.Edit
        
        FR_Table.Fields("[Jumpspeed]") = Jumpspeed
        FR_Table.Fields("[Jumpdelay]") = Jumpdelay
        FR_Table.Fields("[Polygondelay]") = Polygondelay
        FR_Table.Fields("[Markdelay]") = Markdelay
       
        FR_Table.Fields("[Laseroffdelay]") = Laseroffdelay
        FR_Table.Fields("[Laserondelay]") = Laserondelay
        FR_Table.Fields("[Wobblesize]") = Wobblesize
        FR_Table.Fields("[Wobblefrequency]") = Wobblefrequency
        
        If (chkAll.value = vbChecked) Then
                FR_Table.Fields("[Markspeed]") = Markspeed
                FR_Table.Fields("[Laserpower]") = Laserpower
                FR_Table.Fields("[Frequency]") = T1
                FR_Table.Fields("[PulseWidth]") = T2
        End If
        FR_Table.Update
        FR_Table.MoveNext
Loop

MsgBox "SetObjProfiles Complete", vbInformation, "Laser"


End Sub

Private Sub cmdExit_Click()

frmMain.Show

End Sub

Private Sub cmdM1_Click()

txtMarkspeed.Text = Val(txtMarkspeed.Text) - 1
If Val(txtMarkspeed.Text) < 0 Then
    txtMarkspeed.Text = 0
End If

End Sub

Private Sub cmdM2_Click()

txtPulseWidth.Text = Val(txtPulseWidth.Text) - 10
If Val(txtPulseWidth.Text) < 2 Then
    txtPulseWidth.Text = 2
End If

End Sub

Private Sub cmdMM1_Click()

txtMarkspeed.Text = Val(txtMarkspeed.Text) - 1
If Val(txtMarkspeed.Text) < 0 Then
    txtMarkspeed.Text = 0
End If

End Sub

Private Sub cmdMM2_Click()

txtPulseWidth.Text = Val(txtPulseWidth.Text) - 10
If Val(txtPulseWidth.Text) < 2 Then
    txtPulseWidth.Text = 2
End If

End Sub

Private Sub cmdP1_Click()

txtMarkspeed.Text = Val(txtMarkspeed.Text) + 1
If Val(txtMarkspeed.Text) > 30000 Then
    txtMarkspeed.Text = 30000
End If

End Sub

Private Sub cmdP2_Click()

txtPulseWidth.Text = Val(txtPulseWidth.Text) + 1
If Val(txtPulseWidth.Text) > 65535 Then
    txtPulseWidth.Text = 65535
End If

End Sub

Private Sub cmdPP1_Click()

txtMarkspeed.Text = Val(txtMarkspeed.Text) + 10
If Val(txtMarkspeed.Text) > 30000 Then
    txtMarkspeed.Text = 30000
End If

End Sub

Private Sub cmdPP2_Click()

txtPulseWidth.Text = Val(txtPulseWidth.Text) + 10
If Val(txtPulseWidth.Text) > 65535 Then
    txtPulseWidth.Text = 65535
End If

End Sub

Private Sub cmdScale_Click()

'ScaleObj

ObjIndex = CLng(txtObjIndex.Text)

Dim HScale As Double
Dim VScale As Double

Dim XMirror As Long 'VALID VALUES 0,1,2
Dim YMirror As Long

HScale = Format(Val(Text16.Text), "##0.000")
VScale = Format(Val(Text17.Text), "##0.000")

AutomationInterface.ScaleObj ObjIndex, HScale, VScale, XMirror, YMirror


End Sub

Private Sub cmdSetAll_Click()

Dim sSQL As String
Set FR_Database = OpenDatabase(ATC_LASER_BD)
sSQL = "SELECT * FROM [FIXTURE] WHERE [MATRIX ID] =" & MATRIX_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

ProfileIndex = 0

Markspeed = FR_Table.Fields("[Markspeed]")  'inches or bits per second
Jumpspeed = FR_Table.Fields("[Jumpspeed]")  'inches or bits per second

Markspeed_Bits = Format(Markspeed * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")
Jumpspeed_Bits = Format(Jumpspeed * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")

Jumpdelay = FR_Table.Fields("[Jumpdelay]")
Polygondelay = FR_Table.Fields("[Polygondelay]")
Markdelay = FR_Table.Fields("[Markdelay]")
Laserpower = FR_Table.Fields("[Laserpower]")
Laseroffdelay = FR_Table.Fields("[Laseroffdelay]")
Laserondelay = FR_Table.Fields("[Laserondelay]")
Wobblesize = FR_Table.Fields("[Wobblesize]")
Wobblefrequency = FR_Table.Fields("[Wobblefrequency]")
T1 = FR_Table.Fields("[Frequency]")
T2 = FR_Table.Fields("[PulseWidth]")

Dim i As Integer
For i = 0 To ObjectCount - 1
        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        AutomationInterface.SetObjProfile i, ProfileIndex, Markspeed_Bits, Jumpspeed_Bits, Jumpdelay, Markdelay, Polygondelay, Laserpower, Laseroffdelay, Laserondelay, Eightbitword, T1, T2, zAxis, Varijumpdelay, Varijumplength, Wobblesize, Wobblefrequency, Powerreset, Varipolydelay
Next i

MsgBox "SetObjProfiles Complete", vbInformation, "Laser"


End Sub

Private Sub Command1_Click()

'RotateObj

ObjIndex = CLng(txtObjIndex.Text)

Dim Angle As Long
Angle = Val(Text18.Text)

AutomationInterface.RotateObj ObjIndex, Angle

End Sub

Private Sub cmdM3_Click()

txtFrequency.Text = Val(txtFrequency.Text) - 0.01
If Val(txtFrequency.Text) < 0.02 Then
    txtFrequency.Text = 0.02
End If

End Sub

Private Sub cmdP3_Click()

txtFrequency.Text = Val(txtFrequency.Text) + 0.01
If Val(txtFrequency.Text) > 250 Then
    txtFrequency.Text = 250
End If

End Sub

Private Sub cmdPP3_Click()

txtFrequency.Text = Val(txtFrequency.Text) + 10
If Val(txtFrequency.Text) > 250 Then
    txtFrequency.Text = 250
End If

End Sub

Private Sub cmdMM3_Click()

txtFrequency.Text = Val(txtFrequency.Text) - 10
If Val(txtFrequency.Text) < 0.02 Then
    txtFrequency.Text = 0.02
End If

End Sub


Private Sub Form_Load()

Caption = "ATC Alase Technologies - COM Interface Tester Text Profile"

txtObjIndex.Text = ObjIndex

Data1.DatabaseName = ATC_LASER_BD

Dim sSQL As String
sSQL = "SELECT * FROM [FIXTURE] WHERE [MATRIX ID] =" & MATRIX_ID
                                   
Data1.RecordSource = sSQL
Data1.Refresh

lblMatrix_ID.Caption = MATRIX_ID

End Sub

Private Sub Option1_Click()
txtObjIndex.Text = 0
End Sub

Private Sub Option2_Click()
txtObjIndex.Text = 1
End Sub

Private Sub Option3_Click()
txtObjIndex.Text = 2
End Sub

Private Sub Option4_Click()
txtObjIndex.Text = 3
End Sub

Private Sub Option5_Click()
txtObjIndex.Text = 4
End Sub

Private Sub Option6_Click()
txtObjIndex.Text = 5
End Sub
