VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tray Laser Main"
   ClientHeight    =   11160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18810
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11160
   ScaleWidth      =   18810
   Visible         =   0   'False
   Begin VB.OptionButton Option17 
      Caption         =   "[17]  217-1077 F Case  10 X 10 W&&P"
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
      Left            =   360
      TabIndex        =   98
      Top             =   8325
      Width           =   4400
   End
   Begin VB.OptionButton Option16 
      Caption         =   "[16]  217-1077  L Case  10 X 10 W&&P"
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
      Left            =   360
      TabIndex        =   97
      Top             =   7950
      Width           =   4400
   End
   Begin VB.OptionButton Option15 
      Caption         =   "[15]  217-1077  S Case  10 X 10 W&&P"
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
      Left            =   360
      TabIndex        =   96
      Top             =   7575
      Width           =   4400
   End
   Begin VB.CommandButton cmdRefresh6 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   15840
      TabIndex        =   95
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   14760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.OptionButton Option14 
      Caption         =   "[14]  301-H89  E Case 6 X 4 Vertical"
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
      Left            =   360
      TabIndex        =   93
      Top             =   10080
      Width           =   4400
   End
   Begin VB.Data Data4 
      Caption         =   "Data4  FROM [TBL SIZE LOC]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Frame FrameAbrasize 
      Caption         =   "Paraylene Demasking "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      TabIndex        =   72
      Top             =   7080
      Width           =   9495
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE X1"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   8
         Left            =   840
         TabIndex        =   83
         Text            =   "LINE X1"
         ToolTipText     =   "LINE X1"
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE X2"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   82
         Text            =   "LINE X2"
         ToolTipText     =   "LINE X2"
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LEN H"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   4
         Left            =   3840
         TabIndex        =   81
         Text            =   "LEN H"
         ToolTipText     =   "LEN H"
         Top             =   480
         Width           =   540
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE Y1"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   80
         Text            =   "LINE Y1"
         ToolTipText     =   "LINE Y1"
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE Y2"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   79
         Text            =   "LINE Y2"
         ToolTipText     =   "LINE Y2"
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "REP"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   78
         Text            =   "REP"
         ToolTipText     =   "REP"
         Top             =   960
         Width           =   540
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LEN V"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   10
         Left            =   5640
         TabIndex        =   77
         Text            =   "LEN"
         ToolTipText     =   "LEN"
         Top             =   480
         Width           =   540
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "SPACE"
         DataSource      =   "Data4"
         Height          =   285
         Index           =   18
         Left            =   5640
         TabIndex        =   76
         Text            =   "SPACE"
         ToolTipText     =   "SPACE"
         Top             =   960
         Width           =   540
      End
      Begin VB.TextBox txtFrequencyA 
         BackColor       =   &H00FFC0FF&
         DataField       =   "ABRASIZE_Frequency"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   8400
         TabIndex        =   75
         Text            =   "XXX"
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtMarkspeedA 
         BackColor       =   &H00FFC0FF&
         DataField       =   "ABRASIZE_Markspeed"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   8400
         TabIndex        =   74
         Text            =   "XXXXXX"
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox txtPulseWidthA 
         BackColor       =   &H00FFC0FF&
         DataField       =   "ABRASIZE_PulseWidth"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   8400
         TabIndex        =   73
         Text            =   "X"
         ToolTipText     =   "PulseWidth"
         Top             =   1080
         Width           =   800
      End
      Begin VB.Label Label1 
         Caption         =   "Line 1"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   92
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Line 2"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   91
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Length Horiz"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   90
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "REPETITION"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   89
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Length Vert"
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   88
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "SPACE"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   87
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label20 
         Caption         =   "Frequency (kHz)"
         Height          =   255
         Left            =   6720
         TabIndex        =   86
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label21 
         Caption         =   "Mark Speed (in/ms)"
         Height          =   255
         Left            =   6720
         TabIndex        =   85
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label24 
         Caption         =   "Pulse Width (us)"
         Height          =   255
         Left            =   6720
         TabIndex        =   84
         Top             =   1080
         Width           =   1500
      End
   End
   Begin VB.CommandButton CommandExit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exit to Main"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   10680
      Width           =   2000
   End
   Begin VB.OptionButton Option13 
      Caption         =   "[13]  301-476  H Transfer 4 X 3 MS"
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
      Left            =   360
      TabIndex        =   70
      Top             =   9600
      Width           =   4400
   End
   Begin VB.Frame fraWS 
      Caption         =   " Scan W.O."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5160
      TabIndex        =   54
      Top             =   8640
      Width           =   4095
      Begin VB.CommandButton CommandReset 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Reset"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton CommandSet 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Set"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtSQ 
         BackColor       =   &H00FFFFC0&
         DataField       =   "START QTY"
         Height          =   300
         Left            =   1560
         TabIndex        =   60
         Text            =   "12345"
         ToolTipText     =   "Start Qty [START QTY]"
         Top             =   1440
         Width           =   705
      End
      Begin VB.TextBox txtLot 
         BackColor       =   &H00FFFFC0&
         DataField       =   "LOT NUM"
         Height          =   300
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   59
         Text            =   "1234567890"
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox txtWorkOrder 
         BackColor       =   &H00FFFFC0&
         DataField       =   "WORK ORDER"
         Height          =   300
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   58
         Text            =   "123456789012"
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtOrderQty 
         BackColor       =   &H00FFFFC0&
         DataField       =   "QUANTITY"
         Height          =   300
         Left            =   1560
         TabIndex        =   57
         Text            =   "12345"
         ToolTipText     =   "Units Produced"
         Top             =   1800
         Width           =   705
      End
      Begin VB.TextBox txtDefects 
         BackColor       =   &H00C0FFC0&
         DataField       =   "REJECTS"
         Height          =   300
         Left            =   3240
         TabIndex        =   56
         Text            =   "12345"
         ToolTipText     =   "Defects"
         Top             =   1440
         Width           =   585
      End
      Begin VB.TextBox txtATCPart 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ATC PART"
         Height          =   300
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   55
         Text            =   "100E102F"
         ToolTipText     =   "Test for Valid Tolerance on exit field"
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblInfo 
         Caption         =   "Start Qty:"
         Height          =   300
         Index           =   10
         Left            =   240
         TabIndex        =   66
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Caption         =   "Lot Number:"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   65
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "W.O./Lot#:"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   64
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "Quantity:"
         Height          =   300
         Index           =   5
         Left            =   240
         TabIndex        =   63
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Caption         =   "Defects :"
         Height          =   300
         Index           =   6
         Left            =   2400
         TabIndex        =   62
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "ATC Part :"
         Height          =   300
         Index           =   8
         Left            =   240
         TabIndex        =   61
         Top             =   720
         Width           =   1275
      End
   End
   Begin VB.CommandButton CommandInit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Init"
      Height          =   300
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5460
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton CommandBack 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Data Base Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Work Sheet [1]"
      Height          =   300
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Test [0]"
      Height          =   300
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4860
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ATC P/N:"
      Height          =   300
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton Option9 
      Caption         =   "[9]  217-1086 Lica  Case  20 X 20 Chip"
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
      Left            =   360
      TabIndex        =   14
      Top             =   9120
      Width           =   4400
   End
   Begin VB.CommandButton cmdAxis 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Axis 1,2,3"
      Height          =   300
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame FrameTest 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   11280
      TabIndex        =   37
      Top             =   8640
      Width           =   2055
      Begin VB.CommandButton cmdConfiguration 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Configuration"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Width           =   1600
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Power && Font /Logo"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1600
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NI Motion Test"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   1600
      End
      Begin VB.CommandButton cmdConfig 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Carrier Tray Config"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Width           =   1600
      End
   End
   Begin VB.OptionButton Option11 
      Caption         =   "[11] 217-1076 E  2 X 4 Transfer Molded"
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
      Left            =   360
      TabIndex        =   12
      Top             =   6360
      Width           =   4400
   End
   Begin VB.OptionButton Option10 
      Caption         =   "[10] 217-1074 E  2 X 4 Transfer MS"
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
      Left            =   360
      TabIndex        =   11
      Top             =   5925
      Width           =   4400
   End
   Begin VB.OptionButton Option12 
      Caption         =   "[12] 217-1075 C 3 X 4 MS"
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
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   4400
   End
   Begin VB.Frame Frame1 
      Caption         =   " Motion Initialize "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   13440
      TabIndex        =   34
      Top             =   8640
      Width           =   2175
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Load Parameters"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1600
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Axis 1 Tray "
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "FindReverseLimit"
         Top             =   720
         Width           =   1600
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Axis 2 Camera "
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "FindReverseLimit"
         Top             =   1080
         Width           =   1600
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Axis 3 Laser "
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "FindReverseLimit"
         Top             =   1440
         Width           =   1600
      End
   End
   Begin VB.Frame fraText 
      Caption         =   " Default "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   9360
      TabIndex        =   33
      Top             =   8640
      Width           =   1815
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   20
         Text            =   "ABCD"
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   23
         Text            =   "5678"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   22
         Text            =   "EFGH"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   21
         Text            =   "1234"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Label LabelLogo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   360
         TabIndex        =   68
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdTray 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Carrier Tray"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10680
      Width           =   2000
   End
   Begin VB.Data Data2 
      Caption         =   "Data2  FROM [TBL SIZE LOC]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.CommandButton cmdRefresh2 
      Caption         =   "Refresh2"
      Height          =   300
      Left            =   12600
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   12600
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option8 
      Caption         =   "[8]  217-1077  B Case  10 X 10 W&&P"
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
      Left            =   360
      TabIndex        =   13
      Top             =   7200
      Width           =   4400
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5]  301-412 E Case  6 X 4 W&&P"
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
      Left            =   360
      TabIndex        =   8
      Top             =   4440
      Value           =   -1  'True
      Width           =   4400
   End
   Begin VB.OptionButton Option6 
      Caption         =   "[6]  301-412 E Case  4 X 4 Wire RW"
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
      Left            =   360
      TabIndex        =   9
      Top             =   4800
      Width           =   4400
   End
   Begin VB.OptionButton Option7 
      Caption         =   "[7]  301-412 E Case  2 X 4  Lead AR"
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
      Left            =   360
      TabIndex        =   10
      Top             =   5160
      Width           =   4400
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4]  301-414 C Case  3 X 4 Lead AR"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   4400
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3]  301-414 C Case  4 X 4 Wire RW"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   4400
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2]  217-1083 C Case  9 X 4  W&&P"
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
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   4400
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1]  301-413 E  Transfer Molded  2 X 4 Lead Inverted MS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   4515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "XXXX"
      Top             =   120
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1935
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "FROM [TBL Power]"
      Top             =   2040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3413
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "Data1 [TBL Power]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1095
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "FROM  [TBL SIZE LOC]"
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "Data2  [TBL SIZE LOC]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Width           =   4815
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   36
      Top             =   1920
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
      Height          =   3615
      Left            =   15720
      TabIndex        =   94
      ToolTipText     =   "FROM [TBL Power]"
      Top             =   7440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6376
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "Data1 [TBL Power]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   99
      Top             =   6960
      Width           =   4815
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Initialize Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7320
      TabIndex        =   39
      Top             =   5160
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.Label LabelCase 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   16560
      TabIndex        =   53
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LabelLOCATION_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XX"
      Height          =   300
      Left            =   16440
      TabIndex        =   48
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OP_Mode 0 Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14640
      TabIndex        =   47
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "[TBL SIZE LOC]   [TBL Power]"
      Height          =   300
      Left            =   9960
      TabIndex        =   45
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "118 LASER MATRIX.MDB"
      Height          =   300
      Left            =   14640
      TabIndex        =   44
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DPSS Lasers"
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
      Left            =   9840
      TabIndex        =   43
      Top             =   120
      Width           =   4395
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default: "
      Height          =   300
      Left            =   14640
      TabIndex        =   42
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      Height          =   300
      Left            =   16320
      TabIndex        =   41
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
      Height          =   300
      Left            =   14640
      TabIndex        =   40
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v  [3] Carrier Tray Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   32
      Top             =   1320
      Width           =   4395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<< [2] Font Size Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   9840
      TabIndex        =   31
      Top             =   600
      Width           =   4395
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "[1] Carrier Tray Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   30
      Top             =   960
      Width           =   4395
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Top             =   120
      Width           =   4170
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAxis_Click()

FindReverseLimitAxis

End Sub

Private Sub cmdConfig_Click()

frmMain.Hide
frmTray.Show

End Sub

Private Sub cmdConfiguration_Click()

frmMain.Hide
frmConfiguration.Show

End Sub



Private Sub cmdRefresh_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TBL_ID],[SERIES],[CASE],[VALUE],[COATING],[ATC PART]," & _
               "format([Frequency],'0.00')," & _
               "format([Markspeed],'0.000')," & _
               "format([PulseWidth],'0.00'),[ANGLE],[MARK PARA],[GPWR_ID],[Z HEIGHT],format([HEIGHT],'0.000'),[CAMERA POS] " & _
       "FROM [TBL Power] " & _
       "WHERE [ACTIVE] = Yes AND " & _
             "[CASE]='" & CASE_ID & "' AND " & _
             "[TRAY_ID]=" & TRAY_ID & " " & _
       "ORDER BY [SERIES],[COATING],[VALUE]"
                                   
Select Case TRAY_ID
Case 8
        sSQLF = "   |^PID|<Series        ||^DV Range       |<Coating          |<ATC Part                |>Frequency |>Mk Speed|>Pul Wid|^Angle |^PG       |^MST   |Steps |Inch      |>CP      "
Case Else
        sSQLF = "   |^PID|<Series   |^Case|^DV Range       |<Coating          |<ATC Part                |>Frequency |>Mk Speed|>Pul Wid|^Angle |^PG       |^MST   |Steps |Inch      |>CP      "
End Select


Data1.RecordSource = sSQL
Data1.Refresh
 
MSFlexGrid1.FormatString = sSQLF


End Sub

Private Sub cmdRefresh2_Click()

Dim sSQL As String
Dim sSQLF As String
     
sSQL = "SELECT [SIZE_LOC_ID],[CASE NAME]," & _
              "format([FONT HEIGHT],'0.000') " & _
       "FROM [TBL SIZE LOC] WHERE [ACTIVE] = Yes AND [CASE]='" & CASE_ID & "'"
                                         
sSQLF = "   |^SIZE_LOC_ID|<Case Size          |Font Height "

Data2.RecordSource = sSQL
Data2.Refresh
 
MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub cmdRefresh6_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [GPWR_ID],[TRAY_ID],[TBL_ID] " & _
       "FROM [TBL Power] " & _
       "WHERE [ACTIVE] = Yes AND " & _
             "[GPWR_ID] > 0 " & _
       "ORDER BY [GPWR_ID] "
                                   
sSQLF = "   |^MST    |^TRAY_ID|^PID   "
 
Data5.RecordSource = sSQL
Data5.Refresh
 
MSFlexGrid5.FormatString = sSQLF

End Sub

Private Sub cmdTray_Click()

Static bInitialized As Boolean

Select Case INITIALIZE_TRAY
Case 1

        If (bInitialized = False) Then
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            lblMessage.Visible = True
            lblMessage.Caption = "Initialize Controller"
            Initialize_Controller
            lblMessage.Caption = "Disable Home"
            DisableHome
            lblMessage.Caption = "Load Parameters"
            Load_Parameters
            lblMessage.Caption = "Find Limit Axis"
            FindReverseLimitAxis
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            lblMessage.Visible = False
            bInitialized = True
        End If

End Select

If (Option1.value = True) Then
    TRAY_ID = 1
End If

If (Option2.value = True) Then
    TRAY_ID = 2
End If
If (Option3.value = True) Then
    TRAY_ID = 3
End If
If (Option4.value = True) Then
    TRAY_ID = 4
End If

If (Option9.value = True) Then
    TRAY_ID = 9
End If

If (Option5.value = True) Then
    TRAY_ID = 5
End If

If (Option6.value = True) Then
    TRAY_ID = 6
End If
If (Option7.value = True) Then
    TRAY_ID = 7
End If
If (Option10.value = True) Then
    TRAY_ID = 10
End If
If (Option11.value = True) Then
    TRAY_ID = 11
End If

If (Option8.value = True) Then
    TRAY_ID = 8
End If
If (Option15.value = True) Then
    TRAY_ID = 16
End If
If (Option16.value = True) Then
    TRAY_ID = 17
End If
If (Option17.value = True) Then
    TRAY_ID = 18
End If



If (Option12.value = True) Then
    TRAY_ID = 12
End If

If (Option13.value = True) Then
    TRAY_ID = 13
End If

If (Option14.value = True) Then
    TRAY_ID = 14
End If

LASER_TXT1 = UCase(Text1.Text)
LASER_TXT2 = UCase(Text2.Text)
LASER_TXT3 = UCase(Text3.Text)
LASER_TXT4 = UCase(Text4.Text)

Initialize_Fire_Matrix

Select Case TRAY_ID
Case 13, 14
            TRAY_MARK_ANGLE = MARK_ANGLE_ROTATED
Case Else
            TRAY_MARK_ANGLE = MARK_ANGLE_DEFAULT
End Select

'If (Load_Job = 0) Then
       Load_Job_From_File
'End If

Select Case TRAY_ID
Case 14
            TRAY_MARK_ANGLE = MARK_ANGLE_ROTATED
            frm412.Show 'E Case
Case 13
            TRAY_MARK_ANGLE = MARK_ANGLE_ROTATED
            frm103.Show 'H Case
Case 1
            frm413.Show
Case 2, 3, 4, 12
            frm414.Show
Case 5, 6, 7, 10, 11
            frm412.Show
Case 8, 16, 17, 18
            frm10x10.Show
Case 9
            frm20x20.Show
End Select

frmMain.Hide

End Sub

Private Sub Command1_Click()

frmMain.Hide
frmPowerFactors.Show

End Sub

Private Sub Command10_Click()

OP_MODE = 1
Select Case OP_MODE
Case 0
        Label7.Caption = "OP_Mode [0 Test]"
Case 1
        Label7.Caption = "OP_Mode [1 WS]"
End Select

Configuration (FWRITE)

End Sub

Private Sub Command2_Click()

frmMain.Hide
frmMotion.Show

End Sub

Private Sub Command3_Click()

Initialize_Controller

DisableHome

Load_Parameters

MsgBox "NI Motion Parameters Complete", vbInformation, "ATC Tray Laser System"

End Sub

Private Sub Command4_Click()

FindReverseLimit 1

End Sub

Private Sub Command5_Click()

FindReverseLimit 2

End Sub

Private Sub Command7_Click()

FindReverseLimit 3

End Sub

Private Sub Command8_Click()

If ValidPart(txtATCPart.Text) = False Then
        MsgBox "Invalid Part Format", vbInformation, "ATC "
        Exit Sub
End If

SERIES_ID = UCase(Mid(txtATCPart.Text, 1, 4))

Dim sSERIES As String
Select Case SERIES_ID
Case "800E"
            sSERIES = "800"
Case "100E", "710E"
            sSERIES = "100/710"
Case "100C", "710C"
            sSERIES = "100/710"
Case "100B", "710B"
            sSERIES = "100/710"
Case Else
            MsgBox "Invalid Part Series", vbInformation, "ATC "
End Select

If (Mid$(txtATCPart.Text, 6, 1) = "R") Then
            DV_ID = Val(Mid$(txtATCPart.Text, 5, 1) & "." & Mid$(txtATCPart.Text, 7, 1))
Else
            DV_ID = Val(Mid$(txtATCPart.Text, 5, 2) & "E" & Mid$(txtATCPart.Text, 7, 1))
End If

CASE_ID = UCase(Mid(txtATCPart.Text, 4, 1))

LabelCase.Caption = CASE_ID

Text1.Text = UCase(Mid(txtATCPart.Text, 5, 4))

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TBL_ID],[ATC DWG],[SERIES],[CASE],[VALUE],[COATING],[ATC PART] " & _
       "FROM [TBL Power] " & _
       "WHERE [ACTIVE] = Yes AND " & _
             "[CASE]='" & CASE_ID & "' AND " & _
             "[SERIES]='" & sSERIES & "' AND " & _
             "[DV MIN] <=" & DV_ID & " AND [DV MAX] >=" & DV_ID
                                   
sSQLF = "   ||<ATC Drawing                  |<Series    |^Case|^DV Range        |<Coating          |<ATC Part               "

'Data1.RecordSource = sSQL
'Data1.Refresh
 
'MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub Command9_Click()
OP_MODE = 0

Select Case OP_MODE
Case 0
        Label7.Caption = "OP_Mode [0 Test]"
Case 1
        Label7.Caption = "OP_Mode [1 WS]"
End Select

Configuration (FWRITE)

End Sub

Private Sub CommandBack_Click()
On Error GoTo Network_Mode_ErrorAll

Dim dTime As Single
Dim SourceFile As String
Dim DestinationFile As String

Dim FSO As New FileSystemObject

Screen.MousePointer = vbHourglass

Dim i As Integer
Dim Seconds As Long
Dim MinSec As String
 
SourceFile = ATC_LASER_BD

Select Case LOCATION_ID
Case "NY"
            DestinationFile = SERVER_DB_NY & "118 LASER MATRIX.MDB"
Case "JR"
            DestinationFile = SERVER_DB_JR & "118 LASER MATRIX.MDB"
End Select

FSO.CopyFile SourceFile, DestinationFile, True

Seconds = Format(Timer - dTime, "0")
MinSec = Format(Seconds / 60, "00") & ":" & Format(Seconds Mod 60, "00")
                
Screen.MousePointer = vbDefault

MsgBox "Succesful", vbInformation, "ATC DataBase System"

Exit Sub

Network_Mode_ErrorAll:

Screen.MousePointer = vbDefault
MsgBox "Unsuccesful", vbCritical, "ATC DataBase System"
  
Exit Sub
End Sub

Private Sub CommandExit_Click()
Unload Me
End Sub

Private Sub CommandInit_Click()

If INITIALIZE_TRAY = 0 Then
    INITIALIZE_TRAY = 1
Else
    INITIALIZE_TRAY = 0
End If
CommandInit.Caption = "Init " & INITIALIZE_TRAY
Configuration (FWRITE)

End Sub

Private Sub CommandReset_Click()
Option1.BackColor = &H8000000F
Option2.BackColor = &H8000000F
Option3.BackColor = &H8000000F
Option4.BackColor = &H8000000F
Option5.BackColor = &H8000000F
Option6.BackColor = &H8000000F
Option7.BackColor = &H8000000F
Option8.BackColor = &H8000000F
Option9.BackColor = &H8000000F
Option10.BackColor = &H8000000F
Option11.BackColor = &H8000000F
Option12.BackColor = &H8000000F

Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
Option9.Enabled = True
Option10.Enabled = True
Option11.Enabled = True
Option12.Enabled = True
End Sub

Private Sub CommandSet_Click()

Option1.BackColor = &H8000000F
Option2.BackColor = &H8000000F
Option3.BackColor = &H8000000F
Option4.BackColor = &H8000000F
Option5.BackColor = &H8000000F
Option6.BackColor = &H8000000F
Option7.BackColor = &H8000000F
Option8.BackColor = &H8000000F
Option9.BackColor = &H8000000F
Option10.BackColor = &H8000000F
Option11.BackColor = &H8000000F
Option12.BackColor = &H8000000F

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option9.Enabled = False
Option10.Enabled = False
Option11.Enabled = False
Option12.Enabled = False

Dim ValidChar As String
Dim SearchChar As String
Dim MYPOS As Integer
Dim found As Integer

Select Case Mid(ATC_PART_ID, 4, 1)
Case "E"
        
        Option1.value = True
              
        ValidChar = "AM"
        SearchChar = Mid(ATC_PART_ID, 9, 1)
        MYPOS = InStr(1, ValidChar, SearchChar, 1)
        Select Case MYPOS
        Case Is <> 0
                ValidChar = "78"
                SearchChar = Mid(ATC_PART_ID, 10, 1)
                MYPOS = InStr(1, ValidChar, SearchChar, 1)
                Select Case MYPOS
                Case Is <> 0
                        Option1.BackColor = &HFFC0FF
                        Option1.Enabled = True
                        Option1.value = True
                        found = 1
                End Select
        End Select
    
        If (found = 0) Then
                Select Case Mid(ATC_PART_ID, 9, 1)
                Case "Q", "R", "O"
                        Option6.BackColor = &HFFC0FF
                        Option6.value = True
                        Option6.Enabled = True
                Case "C", "I", "P", "S", "T", "W", "Y"
                        Option5.BackColor = &HFFC0FF
                        Option5.value = True
                        Option5.Enabled = True
                Case "A", "G", "J", "U"
                        Option7.BackColor = &HFFC0FF
                        Option7.value = True
                        Option7.Enabled = True
                Case "D", "E", "K", "M"
                        Select Case Mid(ATC_PART_ID, 11, 1)
                        Case "X"
                            Option10.BackColor = &HFFC0FF
                            Option10.value = True
                            Option10.Enabled = True
                        Case Else
                            Option11.BackColor = &HFFC0FF
                            Option11.value = True
                            Option11.Enabled = True
                        End Select
                End Select
        End If
Case "B"
        Option8.value = True
        Option8.BackColor = &HFFC0FF
        Option8.Enabled = True
Case "C"
        Option2.value = True
        Select Case Mid(ATC_PART_ID, 9, 1)
        Case "Q", "R", "O"
                    Option3.BackColor = &HFFC0FF
                    Option3.value = True
                    Option3.Enabled = True
        Case "C", "I", "P", "S", "T", "W", "Y"
                    Option2.BackColor = &HFFC0FF
                    Option2.value = True
                    Option2.Enabled = True
        Case "A", "G", "J", "U"
                    Option4.BackColor = &HFFC0FF
                    Option4.value = True
                    Option4.Enabled = True
        Case "D", "E", "K", "M"
                Select Case Mid(ATC_PART_ID, 11, 1)
                Case "X"
                    Option12.BackColor = &HFFC0FF
                    Option12.value = True
                    Option12.Enabled = True
                Case Else
                    Option4.BackColor = &HFFC0FF
                    Option4.value = True
                    Option4.Enabled = True
                End Select
        End Select
Case Else

End Select

Text1.Text = UCase(Mid(txtATCPart.Text, 5, 4))

ValidChar = "N123579CHB"
SearchChar = Mid(ATC_PART_ID, 10, 1)
MYPOS = InStr(1, ValidChar, SearchChar, 1)
 
Select Case MYPOS
Case 0
    'NOT FOUND     MAG
        LOGO_MODE = LOGO_SIDE
        LabelLogo.Caption = "Mag Side Logo"
Case Else
    'FOUND     NON MAG
        LOGO_MODE = LOGO_ATC
        LabelLogo.Caption = "Non Mag Logo Top ATC"
End Select


End Sub


Private Sub Form_Activate()

Select Case OP_MODE
Case 0
        Label7.Caption = "OP_Mode [0 Test]"
Case 1
        Label7.Caption = "OP_Mode [1 WS]"
End Select

If (Option1.value = True) Then
        Option1_Click
End If
If (Option2.value = True) Then
        Option2_Click
End If
If (Option3.value = True) Then
        Option3_Click
End If
If (Option4.value = True) Then
        Option4_Click
End If
If (Option5.value = True) Then
        Option5_Click
End If
If (Option6.value = True) Then
        Option6_Click
End If
If (Option7.value = True) Then
        Option7_Click
End If
If (Option8.value = True) Then
        Option8_Click
End If
If (Option9.value = True) Then
        Option9_Click
End If
If (Option10.value = True) Then
        Option10_Click
End If
If (Option11.value = True) Then
        Option11_Click
End If
If (Option12.value = True) Then
        Option12_Click
End If

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Form_Load()

Caption = "Tray Laser Main      " & ATC_DWG & "         " & ATC_VERSION
 
lblDate.Caption = Date
lblUser.Caption = strComputerName
lblIP.Caption = IP_ADDRESS
LabelLOCATION_ID.Caption = LOCATION_ID
 
Data1.DatabaseName = ATC_LASER_BD
Data2.DatabaseName = ATC_LASER_BD
Data3.DatabaseName = ATC_LASER_BD
Data4.DatabaseName = ATC_LASER_BD
Data5.DatabaseName = ATC_LASER_BD

'MSFlexGrid1.Top = 0
'MSFlexGrid1.Left = 0
MSFlexGrid1.Width = 13600
MSFlexGrid1.Height = 5000

Option1_Click

frmMain.Top = FORM_LOC_Y
frmMain.Left = FORM_LOC_X

CommandInit.Caption = "Init " & INITIALIZE_TRAY

txtWorkOrder.Text = ""
txtATCPart.Text = ""
txtLot.Text = ""
txtSQ.Text = "0"
txtOrderQty.Text = "0"
txtDefects.Text = "0"

cmdRefresh6_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

frmOPScreen.Show

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sBuff As String
sBuff = UCase(txtPassword.Text)

If (Button = 2 And Shift = 1) Then
       
    Select Case sBuff
    Case "MIKE" & Mid(Format(Date, "ddd"), 1, 1), "ERIK" & Mid(Format(Date, "ddd"), 1, 1)
                 
                 
    Case "ERIK"
    
    Case Else
             
    End Select
       
Else

    If (CommandBack.Visible = True) Then
            CommandBack.Visible = False
            Command10.Visible = False
            Command9.Visible = False
            CommandInit.Visible = False
            FrameTest.Enabled = False
    Else
            CommandBack.Visible = True
            Command10.Visible = True
            Command9.Visible = True
            CommandInit.Visible = True
            FrameTest.Enabled = True
    End If
    
End If

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
POWER_ID = Val(MSFlexGrid1.Text)
 
Dim sSQL As String
 
sSQL = "SELECT  * " & _
       "FROM [TBL Power] " & _
       "WHERE [TBL_ID] = " & POWER_ID
                                  
Data3.RecordSource = sSQL
Data3.Refresh
 
 
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10
 
End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
SIZE_LOC_ID = Val(MSFlexGrid2.Text)
 
Dim sSQL As String
 
sSQL = "SELECT * " & _
       "FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID
                                  
Data4.RecordSource = sSQL
Data4.Refresh
  
 
 
MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1 '10

End Sub

Private Sub Option1_Click()

CASE_ID = "E"
TRAY_ID = 1

ATC_DWG_ID = "301 - 413"

cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option10_Click()

CASE_ID = "E"
TRAY_ID = 10
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID


ATC_DWG_ID = "217-1074"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option11_Click()

CASE_ID = "E"
TRAY_ID = 11

cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "217-1076"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option12_Click()

CASE_ID = "C"
TRAY_ID = 12
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "217 - 1075"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option13_Click()

CASE_ID = "H"
TRAY_ID = 13
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301 - 476"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option14_Click()

CASE_ID = "E"
TRAY_ID = 14
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301 - H89"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click


End Sub

Private Sub Option15_Click()
CASE_ID = "S"
TRAY_ID = 16
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "217- 1067"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click
End Sub

Private Sub Option16_Click()
CASE_ID = "L"
TRAY_ID = 17
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "217- 1067"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click
End Sub

Private Sub Option17_Click()

CASE_ID = "F"
TRAY_ID = 18
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "217- 1067"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option2_Click()

CASE_ID = "C"

TRAY_ID = 2
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301 - 414"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option3_Click()

CASE_ID = "C"
TRAY_ID = 3
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301 - 414"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option4_Click()

CASE_ID = "C"
TRAY_ID = 4
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301 - 414"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option5_Click()

CASE_ID = "E"

TRAY_ID = 5
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301-412"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option6_Click()

CASE_ID = "E"
TRAY_ID = 6
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301-412"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option7_Click()

CASE_ID = "E"
TRAY_ID = 7
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "301-412"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option8_Click()

CASE_ID = "B"
TRAY_ID = 8
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "217- 1067"

cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub Option9_Click()

CASE_ID = "P"
TRAY_ID = 9
cmdTray.Caption = "Carrier TRAY_ID " & TRAY_ID

ATC_DWG_ID = "LICA"


cmdRefresh_Click
MSFlexGrid1_Click
cmdRefresh2_Click
MSFlexGrid2_Click

End Sub

Private Sub txtATCPart_GotFocus()

txtATCPart.SelStart = 0
txtATCPart.SelLength = Len(txtATCPart)

End Sub

Private Sub txtATCPart_LostFocus()

txtATCPart = UCase(txtATCPart)

Select Case 0
Case 0
    
        ValidPart (txtATCPart.Text)
        
        ATC_PART_ID = txtATCPart.Text
        
        TOLERANCE_ID = "X"
        If (Len(txtATCPart) >= 8) Then
             TOLERANCE_ID = Mid$(txtATCPart, 8, 1)
        Else
        
        End If
        
        If Len(ATC_PART_ID) > 7 Then
        
            Select Case CODE_ID
            Case 0       'WORK ORDERS
                        Select Case TOLERANCE_ID
                        Case "A", "B", "C", "D", "F", "G", "J", "K", "M", "N"
                            ' VALID TOLERANCES
                        Case Else
                                frmValid.Show vbModal
                                txtATCPart = Mid$(txtATCPart, 1, 7) & TOLERANCE_ID
                        End Select
            Case Else
                        'No Tolerance Needed only Design Value
            End Select
        End If
        
        Dim ValidChar As String
        ValidChar = "N123579CHB"
        Dim SearchChar As String
        SearchChar = Mid(ATC_PART_ID, 10, 1)
        
        Dim MYPOS As Integer
        MYPOS = InStr(1, ValidChar, SearchChar, 1)
         
        Select Case MYPOS
        Case 0
            'NOT FOUND     MAG
                LOGO_MODE = LOGO_SIDE
               ' LabelLogo.Caption = "Mag Side Logo"
        Case Else
            'FOUND     NON MAG
                LOGO_MODE = LOGO_ATC
              '  LabelLogo.Caption = "Non Mag Logo Top ATC"
        End Select
    
End Select

End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtWorkOrder_GotFocus()
txtWorkOrder.SelStart = 0
txtWorkOrder.SelLength = Len(txtWorkOrder)
End Sub

Private Sub txtWorkOrder_LostFocus()
 
txtWorkOrder.Text = Trim(txtWorkOrder.Text)
txtWorkOrder.Text = Mid(UCase(txtWorkOrder.Text), 1, 12)

'
'   SCHEDULE LOOKUP BY WORK ORDER
'
 
Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)

If (Val(txtOrderQty.Text) = 0) Then

    Dim sSQL As String
    sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='" & txtWorkOrder.Text & "'"
                  
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
    If (FR_Table.RecordCount <> 0) Then
        If (FR_Table.Fields("[ATC PART]") <> vbNull) Then
            txtATCPart.Text = FR_Table.Fields("[ATC PART]")
            ATC_PART_ID = FR_Table.Fields("[ATC PART]")
        End If
        If (FR_Table.Fields("[LOT NUM]") <> vbNull) Then
            txtLot.Text = FR_Table.Fields("[LOT NUM]")
        End If
        If (FR_Table.Fields("[START QTY]") <> vbNull) Then
            txtOrderQty.Text = FR_Table.Fields("[START QTY]")
        Else
            txtOrderQty.Text = 0
        End If
    End If
    FR_Table.Close
    FR_Database.Close

End If

End Sub
