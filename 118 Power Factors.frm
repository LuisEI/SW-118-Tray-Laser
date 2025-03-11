VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPowerFactors 
   Caption         =   "118 DPSS Power & Scale Factors"
   ClientHeight    =   11880
   ClientLeft      =   4530
   ClientTop       =   675
   ClientWidth     =   18435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11880
   ScaleWidth      =   18435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlus1 
      BackColor       =   &H0080FFFF&
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
      Height          =   360
      Left            =   12855
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdMinus1 
      BackColor       =   &H0080FFFF&
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
      Height          =   360
      Left            =   13230
      Style           =   1  'Graphical
      TabIndex        =   168
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdMinus11 
      BackColor       =   &H0080FFFF&
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
      Height          =   360
      Left            =   13605
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus11 
      BackColor       =   &H0080FFFF&
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
      Height          =   360
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtZHT 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      DataField       =   "Z HEIGHT"
      DataSource      =   "Data2"
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
      Height          =   360
      Left            =   11040
      MaxLength       =   10
      TabIndex        =   165
      Text            =   "Z"
      ToolTipText     =   "Z HEIGHT"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdZHT 
      BackColor       =   &H0080FFFF&
      Caption         =   "To ZHT"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   1680
      Width           =   1400
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Rev Limit Z"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   1680
      Width           =   1400
   End
   Begin VB.CheckBox CheckROTATE 
      Caption         =   "Rotated"
      Height          =   375
      Left            =   5280
      TabIndex        =   160
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame FrameAbrasize 
      Caption         =   " Paraylene Demasking "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      TabIndex        =   127
      Top             =   10560
      Width           =   10575
      Begin VB.TextBox txtPulseWidthA 
         BackColor       =   &H00FFC0FF&
         DataField       =   "ABRASIZE_PulseWidth"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   8760
         TabIndex        =   10
         Text            =   "X"
         ToolTipText     =   "PulseWidth"
         Top             =   1080
         Width           =   800
      End
      Begin VB.TextBox txtMarkspeedA 
         BackColor       =   &H00FFC0FF&
         DataField       =   "ABRASIZE_Markspeed"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   8760
         TabIndex        =   9
         Text            =   "XXXXXX"
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox txtFrequencyA 
         BackColor       =   &H00FFC0FF&
         DataField       =   "ABRASIZE_Frequency"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   8760
         TabIndex        =   8
         Text            =   "XXX"
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "SPACE"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   18
         Left            =   5640
         TabIndex        =   7
         Text            =   "SPACE"
         ToolTipText     =   "SPACE"
         Top             =   960
         Width           =   540
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LEN V"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   10
         Left            =   5640
         TabIndex        =   3
         Text            =   "LEN"
         ToolTipText     =   "LEN"
         Top             =   480
         Width           =   540
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "REP"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   6
         Text            =   "REP"
         ToolTipText     =   "REP"
         Top             =   960
         Width           =   540
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE Y2"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Text            =   "LINE Y2"
         ToolTipText     =   "LINE Y2"
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE Y1"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   1
         Text            =   "LINE Y1"
         ToolTipText     =   "LINE Y1"
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LEN H"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   4
         Left            =   3840
         TabIndex        =   2
         Text            =   "LEN H"
         ToolTipText     =   "LEN H"
         Top             =   480
         Width           =   540
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE X2"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   4
         Text            =   "LINE X2"
         ToolTipText     =   "LINE X2"
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LINE X1"
         DataSource      =   "Data3"
         Height          =   285
         Index           =   8
         Left            =   840
         TabIndex        =   0
         Text            =   "LINE X1"
         ToolTipText     =   "LINE X1"
         Top             =   480
         Width           =   800
      End
      Begin VB.Label Label24 
         Caption         =   "Pulse Width (us)"
         Height          =   255
         Left            =   6720
         TabIndex        =   159
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label21 
         Caption         =   "Mark Speed (in/ms)"
         Height          =   255
         Left            =   6720
         TabIndex        =   158
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label20 
         Caption         =   "Frequency (kHz)"
         Height          =   255
         Left            =   6720
         TabIndex        =   157
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "SPACE"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   133
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Length Vert"
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   132
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "REPETITION"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   131
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Length Horiz"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   130
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Line 2"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   129
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Line 1"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   128
         Top             =   480
         Width           =   795
      End
   End
   Begin VB.Frame fraLogo 
      Caption         =   " Logo Mode "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   117
      Top             =   4920
      Width           =   2055
      Begin VB.OptionButton optLogo5 
         Caption         =   "[5] Abrasive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   119
         Top             =   1860
         Width           =   1380
      End
      Begin VB.OptionButton optLogo 
         Caption         =   "[6] Serialization"
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
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   118
         Top             =   2235
         Width           =   1740
      End
      Begin VB.OptionButton optLogo4 
         Caption         =   "[4] Omit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1485
         Width           =   1260
      End
      Begin VB.OptionButton optLogo1 
         Caption         =   "[1] Side"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1110
         Width           =   1260
      End
      Begin VB.OptionButton optLogo3 
         Caption         =   "[3]Top"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   735
         Width           =   900
      End
      Begin VB.OptionButton optLogo2 
         Caption         =   "[2] ATC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Load Job"
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
      Left            =   5160
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5760
      Width           =   2600
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      DataField       =   "SERIES"
      DataSource      =   "Data2"
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
      Index           =   17
      Left            =   240
      TabIndex        =   114
      Text            =   "Series"
      ToolTipText     =   "SERIES"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      DataField       =   "COATING"
      DataSource      =   "Data2"
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
      Index           =   19
      Left            =   2520
      TabIndex        =   113
      Text            =   "COATING"
      ToolTipText     =   "COATING"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      DataField       =   "ATC PART"
      DataSource      =   "Data2"
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
      Index           =   21
      Left            =   4440
      TabIndex        =   112
      Text            =   "ATC PART"
      ToolTipText     =   "ATC PART"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Data Data4 
      Caption         =   "Data4  FROM [TBL POWER]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Power"
      Top             =   10920
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.CommandButton CommandRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   10200
      TabIndex        =   111
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   480
      TabIndex        =   110
      Top             =   9480
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Data Data1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data1  FROM [TBL SIZE LOC]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\106 MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL SIZE LOC"
      Top             =   8760
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   350
      Left            =   16320
      TabIndex        =   69
      Top             =   11640
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   " Object Position "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   8400
      TabIndex        =   92
      Top             =   5880
      Width           =   7575
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R LOGO LY TOP"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         TabIndex        =   147
         Text            =   "0.000"
         ToolTipText     =   "LOGO LY TOP"
         Top             =   1380
         Width           =   900
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R LOGO LX TOP"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         TabIndex        =   146
         Text            =   "-0.000"
         ToolTipText     =   "LOGO LX TOP"
         Top             =   1380
         Width           =   900
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R LOGO LX SIDE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         TabIndex        =   145
         Text            =   "-0.000"
         ToolTipText     =   "LOGO LX SIDE"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R LOGO LY SIDE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         TabIndex        =   144
         Text            =   "-0.000"
         ToolTipText     =   "LOGO LY SIDE"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R ATC  X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   30
         Left            =   5280
         TabIndex        =   143
         Text            =   "ATC  X"
         ToolTipText     =   "ATC  X"
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT1 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   29
         Left            =   5280
         TabIndex        =   142
         Text            =   "TXT1 X"
         ToolTipText     =   "TXT1 X"
         Top             =   2520
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT2 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   28
         Left            =   5280
         TabIndex        =   141
         Text            =   "TXT2 X"
         ToolTipText     =   "TXT2 X"
         Top             =   2940
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT3 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   27
         Left            =   5280
         TabIndex        =   140
         Text            =   "TXT3 X"
         ToolTipText     =   "TXT3 X"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT4 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   26
         Left            =   5280
         TabIndex        =   139
         Text            =   "TXT4 X"
         ToolTipText     =   "TXT4 X"
         Top             =   3780
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R ATC  Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   25
         Left            =   6360
         TabIndex        =   138
         Text            =   "ATC  Y"
         ToolTipText     =   "ATC  Y"
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT1 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   24
         Left            =   6360
         TabIndex        =   137
         Text            =   "TXT1 Y"
         ToolTipText     =   "TXT1 Y"
         Top             =   2520
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT2 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   23
         Left            =   6360
         TabIndex        =   136
         Text            =   "TXT2 Y"
         ToolTipText     =   "TXT2 Y"
         Top             =   2940
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT3 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   22
         Left            =   6360
         TabIndex        =   135
         Text            =   "TXT3 Y"
         ToolTipText     =   "TXT3 Y"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0FF&
         DataField       =   "R TXT4 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   20
         Left            =   6360
         TabIndex        =   134
         Text            =   "TXT4 Y"
         ToolTipText     =   "TXT4 Y"
         Top             =   3780
         Width           =   900
      End
      Begin VB.CommandButton cmdUpdate2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Update Record"
         Height          =   300
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "FROM [TBL SIZE LOC] WHERE [CASE SIZE]"
         Top             =   4200
         Width           =   1485
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT4 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   16
         Left            =   3840
         TabIndex        =   67
         Text            =   "TXT4 Y"
         ToolTipText     =   "TXT4 Y"
         Top             =   3780
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT3 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   14
         Left            =   3840
         TabIndex        =   65
         Text            =   "TXT3 Y"
         ToolTipText     =   "TXT3 Y"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT2 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   9
         Left            =   3840
         TabIndex        =   63
         Text            =   "TXT2 Y"
         ToolTipText     =   "TXT2 Y"
         Top             =   2940
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT1 Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   3840
         TabIndex        =   61
         Text            =   "TXT1 Y"
         ToolTipText     =   "TXT1 Y"
         Top             =   2520
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "ATC  Y"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   3840
         TabIndex        =   55
         Text            =   "ATC  Y"
         ToolTipText     =   "ATC  Y"
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT4 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   13
         Left            =   2760
         TabIndex        =   66
         Text            =   "TXT4 X"
         ToolTipText     =   "TXT4 X"
         Top             =   3780
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT3 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   12
         Left            =   2760
         TabIndex        =   64
         Text            =   "TXT3 X"
         ToolTipText     =   "TXT3 X"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT2 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   11
         Left            =   2760
         TabIndex        =   62
         Text            =   "TXT2 X"
         ToolTipText     =   "TXT2 X"
         Top             =   2940
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TXT1 X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2760
         TabIndex        =   60
         Text            =   "TXT1 X"
         ToolTipText     =   "TXT1 X"
         Top             =   2520
         Width           =   900
      End
      Begin VB.TextBox txtScale 
         BackColor       =   &H00FFC0C0&
         DataField       =   "ATC  X"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   15
         Left            =   2760
         TabIndex        =   54
         Text            =   "ATC  X"
         ToolTipText     =   "ATC  X"
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LOGO LY SIDE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         TabIndex        =   57
         Text            =   "-0.000"
         ToolTipText     =   "LOGO LY SIDE"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LOGO LX SIDE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         TabIndex        =   56
         Text            =   "-0.000"
         ToolTipText     =   "LOGO LX SIDE"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LOGO LX TOP"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         TabIndex        =   58
         Text            =   "-0.000"
         ToolTipText     =   "LOGO LX TOP"
         Top             =   1380
         Width           =   900
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00FFC0C0&
         DataField       =   "LOGO LY TOP"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         TabIndex        =   59
         Text            =   "0.000"
         ToolTipText     =   "LOGO LY TOP"
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Rotated Locations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   5280
         TabIndex        =   148
         Top             =   600
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =  SIZE_LOC_ID"
         Height          =   195
         Left            =   360
         TabIndex        =   126
         Top             =   360
         Width           =   4725
      End
      Begin VB.Label Label1 
         Caption         =   "Text 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   360
         TabIndex        =   101
         Top             =   3645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Text 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   360
         TabIndex        =   100
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Text 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   360
         TabIndex        =   99
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Text 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   360
         TabIndex        =   98
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "ATC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   360
         TabIndex        =   97
         ToolTipText     =   "ATC Non Mag"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Loc Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3840
         TabIndex        =   96
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Loc X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   2760
         TabIndex        =   95
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label17 
         Caption         =   "LOGO Side"
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
         Left            =   360
         TabIndex        =   94
         Top             =   1755
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "LOGO Top"
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
         Left            =   360
         TabIndex        =   93
         Top             =   1335
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Object Profile "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   85
      Top             =   2040
      Width           =   6975
      Begin VB.CheckBox Check2 
         Caption         =   "Active"
         DataField       =   "ACTIVE"
         DataSource      =   "Data2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   125
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Update Record"
         Height          =   300
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "[FIXTURE] WHERE [MATRIX ID]"
         Top             =   2280
         Width           =   1485
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
         Left            =   4395
         TabIndex        =   18
         Top             =   1200
         Width           =   800
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
         Left            =   5190
         TabIndex        =   19
         Top             =   1200
         Width           =   800
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
         Left            =   5985
         TabIndex        =   20
         Top             =   1200
         Width           =   800
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
         Left            =   3600
         TabIndex        =   17
         Top             =   1200
         Width           =   800
      End
      Begin VB.TextBox txtFrequency 
         BackColor       =   &H00FFC0FF&
         DataField       =   "Frequency"
         DataSource      =   "Data2"
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
         TabIndex        =   11
         Text            =   "XXX"
         Top             =   600
         Width           =   1095
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
         Left            =   4395
         TabIndex        =   13
         Top             =   600
         Width           =   800
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
         Left            =   5190
         TabIndex        =   14
         Top             =   600
         Width           =   800
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
         Left            =   5985
         TabIndex        =   15
         Top             =   600
         Width           =   800
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
         Left            =   3600
         TabIndex        =   12
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox txtMarkspeed 
         BackColor       =   &H00FFC0FF&
         DataField       =   "Markspeed"
         DataSource      =   "Data2"
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
         TabIndex        =   16
         Text            =   "XXXXXX"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtPulseWidth 
         BackColor       =   &H00FFC0FF&
         DataField       =   "PulseWidth"
         DataSource      =   "Data2"
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
         TabIndex        =   21
         Text            =   "X"
         ToolTipText     =   "PulseWidth"
         Top             =   1800
         Width           =   1095
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
         Left            =   4395
         TabIndex        =   23
         Top             =   1800
         Width           =   800
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
         Left            =   5190
         TabIndex        =   24
         Top             =   1800
         Width           =   800
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
         Left            =   5985
         TabIndex        =   25
         Top             =   1800
         Width           =   800
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
         Left            =   3600
         TabIndex        =   22
         Top             =   1800
         Width           =   800
      End
      Begin VB.Label Label6 
         Caption         =   "POWER_ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   124
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         DataField       =   "TBL_ID"
         DataSource      =   "Data2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   18
         Left            =   2280
         TabIndex        =   123
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label33 
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
         Left            =   240
         TabIndex        =   91
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "[0.02 to 250.0]"
         Height          =   300
         Left            =   2280
         TabIndex        =   90
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label31 
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
         Left            =   240
         TabIndex        =   89
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label30 
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
         Left            =   240
         TabIndex        =   88
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "[0 to 30000]"
         Height          =   300
         Left            =   2280
         TabIndex        =   87
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "[2 to 65535]"
         Height          =   300
         Left            =   2280
         TabIndex        =   86
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSetObjSize 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SetObjSize"
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
      Left            =   5160
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6720
      Width           =   2600
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   84
      Text            =   "XXXX"
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdSetObjProfile 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SetObjProfile,SetObjPos"
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
      Left            =   5160
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7320
      Width           =   2600
   End
   Begin VB.Data Data3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Data3  FROM [TBL SIZE LOC]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\118 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL SIZE LOC"
      Top             =   11880
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.TextBox txtOFFY 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6840
      TabIndex        =   45
      Text            =   "0.000"
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox txtOFFX 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5880
      TabIndex        =   44
      Text            =   "0.000"
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton cmdFireDPPS 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MarkObj Location (X,Y)"
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
      Left            =   5160
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7920
      Width           =   2600
   End
   Begin VB.CommandButton cmdSetObjCharString 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SetObjCharString"
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
      Left            =   5160
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6240
      Width           =   2600
   End
   Begin VB.Frame fraText 
      Caption         =   " Number of Lines "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2520
      TabIndex        =   81
      Top             =   4920
      Width           =   2295
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
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
         Left            =   1080
         TabIndex        =   31
         Text            =   "AAAA"
         Top             =   360
         Width           =   885
      End
      Begin VB.OptionButton OptLine 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   36
         Top             =   1680
         Width           =   645
      End
      Begin VB.OptionButton OptLine 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   34
         Top             =   1240
         Width           =   645
      End
      Begin VB.OptionButton OptLine 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   800
         Width           =   645
      End
      Begin VB.OptionButton OptLine 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
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
         Left            =   1080
         TabIndex        =   37
         Text            =   "DDDD"
         Top             =   1680
         Width           =   885
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
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
         Left            =   1080
         TabIndex        =   35
         Text            =   "CCCC"
         Top             =   1240
         Width           =   885
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFC0&
         DataField       =   " "
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
         Left            =   1080
         TabIndex        =   33
         Text            =   "BBBB"
         Top             =   800
         Width           =   885
      End
   End
   Begin VB.Frame fra3 
      Caption         =   "Text Object "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9480
      TabIndex        =   76
      Top             =   3000
      Width           =   6495
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0E0FF&
         DataField       =   "TEXT XSIZE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   77
         ToolTipText     =   "Width (Horizontal)"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0E0FF&
         DataField       =   "TEXT YSIZE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   47
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtYScale 
         BackColor       =   &H00C0E0FF&
         DataField       =   "TEXT YSCALE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5160
         TabIndex        =   49
         ToolTipText     =   "TEXT YSCALE"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtXScale 
         BackColor       =   &H00C0E0FF&
         DataField       =   "TEXT XSCALE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4080
         TabIndex        =   48
         ToolTipText     =   "TEXT XSCALE"
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "[Vert]"
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
         Left            =   5160
         TabIndex        =   107
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "[Horiz]"
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
         Left            =   4080
         TabIndex        =   106
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "[Vert]"
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
         Left            =   2160
         TabIndex        =   103
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "[Horiz]"
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
         TabIndex        =   102
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label13 
         Caption         =   "Size"
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
         Left            =   360
         TabIndex        =   79
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label12 
         Caption         =   "Scale"
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
         Left            =   3240
         TabIndex        =   78
         Top             =   720
         Width           =   675
      End
   End
   Begin VB.Frame fra4 
      Caption         =   " Logo Object "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9480
      TabIndex        =   74
      Top             =   4440
      Width           =   6495
      Begin VB.TextBox Text21 
         BackColor       =   &H00C0E0FF&
         DataField       =   "LOGO XSIZE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   50
         ToolTipText     =   "LOGO XSIZE"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00C0E0FF&
         DataField       =   "LOGO YSIZE"
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   51
         ToolTipText     =   "LOGO YSIZE"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00C0E0FF&
         DataField       =   "LOGO XSCALE"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4080
         TabIndex        =   52
         Text            =   "0.000"
         ToolTipText     =   "LOGO XSCALE"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0E0FF&
         DataField       =   "LOGO YSCALE"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5160
         TabIndex        =   53
         Text            =   "0.000"
         ToolTipText     =   "LOGO YSCALE"
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "[Vert]"
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
         Left            =   5160
         TabIndex        =   109
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "[Horiz]"
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
         Left            =   4080
         TabIndex        =   108
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "[Vert]"
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
         Left            =   2160
         TabIndex        =   105
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "[Horiz]"
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
         TabIndex        =   104
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label15 
         Caption         =   "Size"
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
         Left            =   360
         TabIndex        =   80
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label27 
         Caption         =   "Scale"
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
         Left            =   3240
         TabIndex        =   75
         Top             =   720
         Width           =   795
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Serialization "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14760
      TabIndex        =   70
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtStartNo 
         Height          =   300
         Left            =   1920
         TabIndex        =   71
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Note : Uses Location Text1"
         Height          =   300
         Index           =   0
         Left            =   480
         TabIndex        =   83
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Start Number:"
         Height          =   300
         Index           =   8
         Left            =   480
         TabIndex        =   72
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Data Data2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Data2 FROM [TBL POWER]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\118 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL POWER"
      Top             =   11400
      Visible         =   0   'False
      Width           =   4020
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   855
      Left            =   7800
      TabIndex        =   46
      ToolTipText     =   "FROM [TBL Power] WHERE [CASE] and [TRAY_ID]"
      Top             =   480
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1508
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "FROM [TBL POWER] WHERE [TBL_ID]"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   240
      TabIndex        =   38
      Top             =   7920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6588
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Frame FrameRotation 
      Caption         =   "  Rotation "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3360
      TabIndex        =   149
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton CommandRotateObj 
         BackColor       =   &H00C0FFC0&
         Caption         =   "RotateObj Angle"
         Height          =   300
         Left            =   360
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   "Rotates object about its center"
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox TextAngle 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   2040
         TabIndex        =   153
         Text            =   "90"
         Top             =   960
         Width           =   500
      End
      Begin VB.TextBox TextY 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   2040
         TabIndex        =   152
         Text            =   "0.00"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox TextX 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   1320
         TabIndex        =   151
         Text            =   "0.00"
         Top             =   1920
         Width           =   500
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "RotateObjEx Angle (X,Y)"
         Height          =   300
         Left            =   360
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   "Rotates object about a coordinate center"
         Top             =   1440
         Width           =   2000
      End
      Begin VB.Label Label7 
         Caption         =   " (-360 to 360)"
         Height          =   300
         Left            =   960
         TabIndex        =   156
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Angle"
         Height          =   300
         Left            =   480
         TabIndex        =   155
         Top             =   960
         Width           =   435
      End
   End
   Begin VB.Label LabelTRAY_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6840
      TabIndex        =   173
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label37 
      Caption         =   "TRAY_ID:"
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
      Left            =   5280
      TabIndex        =   172
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LabelPOWER_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6840
      TabIndex        =   171
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label36 
      Caption         =   "POWER_ID:"
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
      Left            =   5280
      TabIndex        =   170
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LabelMARK_ANGLE 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   6720
      TabIndex        =   162
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label Label26 
      Caption         =   "MARK_ANGLE:"
      Height          =   300
      Left            =   5160
      TabIndex        =   161
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   3300
      Left            =   15480
      Top             =   5880
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "[TBL POWER] WHERE [TBL_ID] = POWER_ID"
      Height          =   255
      Left            =   5040
      TabIndex        =   121
      ToolTipText     =   "CASE NAME"
      Top             =   4920
      Width           =   3525
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "FROM [TBL POWER] WHERE [CASE]  =  CASE and [TRAY_ID] = TRAY_ID"
      Height          =   195
      Left            =   8280
      TabIndex        =   120
      Top             =   120
      Width           =   5475
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   7740
      Left            =   16200
      Top             =   3120
      Width           =   1980
   End
   Begin VB.Label LabelSIZE_LOC_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SIZE_LOC_ID"
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
      Left            =   6840
      TabIndex        =   116
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "SIZE_LOC_ID:"
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
      Left            =   5280
      TabIndex        =   115
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label22 
      Caption         =   "(X,Y)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      TabIndex        =   82
      Top             =   8520
      Width           =   555
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Top             =   0
      Width           =   4170
   End
   Begin VB.Label lblMatrix 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Matrix Caption"
      DataField       =   "CAPTION"
      DataSource      =   "Data2"
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
      Left            =   240
      TabIndex        =   73
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "frmPowerFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()

Data2.UpdateRecord
Data3.UpdateRecord

SERIAL_START_NUMBER = Val(txtStartNo.Text)

Unload Me

End Sub

Private Sub cmdFireDPPS_Click()

cmdFireDPPS.Enabled = False

TRAY_X_OFFSET = Val(txtOFFX.Text)
TRAY_Y_OFFSET = Val(txtOFFY.Text)
   
If (optLogo2.value = True) Then
    LOGO_MODE = LOGO_ATC
End If
If (optLogo3.value = True) Then
    LOGO_MODE = LOGO_TOP
End If
If (optLogo1.value = True) Then
    LOGO_MODE = LOGO_SIDE
End If
If (optLogo4.value = True) Then
    LOGO_MODE = LOGO_OMIT
End If
'ABRASIZE
If (optLogo5.value = True) Then
    LOGO_MODE = LOGO_ABRASIVE
End If
    
If (OptLine(0).value = True) Then
     MARK_MODE = 1
End If
If (OptLine(1).value = True) Then
      MARK_MODE = 2
End If
If (OptLine(2).value = True) Then
      MARK_MODE = 3
End If
If (OptLine(3).value = True) Then
      MARK_MODE = 4
End If
    
Dim HPosition As Long
Dim VPosition As Long
Dim Angle As Long
 
Dim j As Integer
Dim i As Integer
                                                         
For j = 0 To ObjectCount - 1

        '[1] Get Object Position    (HPosition,VPosition)
        '[2] Fire On Object
        'chg 08/16/12
        HPosition = Format((ObjHPosition(j) + TRAY_X_OFFSET) * INCHES_TO_BITS, "0")
        VPosition = Format((ObjVPosition(j) + TRAY_Y_OFFSET) * INCHES_TO_BITS, "0")

        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        
        Select Case j
        Case 0
                Select Case MARK_MODE
                Case 1, 2, 3, 4
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
        Case 1
                Select Case MARK_MODE
                Case 2, 3, 4
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
        Case 2
                Select Case MARK_MODE
                Case 3, 4
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
        Case 3
                Select Case MARK_MODE
                Case 4
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
        Case 4
                Select Case LOGO_MODE
                Case LOGO_SIDE, LOGO_TOP
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
        Case 5
                Select Case LOGO_MODE
                Case LOGO_ATC
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
        Case 6, 7, 8, 9, 10, 11
                'ABRASIVE LINE 1/2/3/4/5/6
                Select Case LOGO_MODE
                Case LOGO_ABRASIVE
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        For i = 1 To REP_ABRASIVE
                            BusyFlag2 = 1
                            While BusyFlag2 = 1
                                AutomationInterface.GetBusyStatus 0, BusyFlag2
                            Wend
                            
                            AutomationInterface.MarkObj j, 0
                            
                        Next i
                End Select
        End Select
        
Next j
                  
'MsgBox "Complete", vbInformation, "Laser"

cmdFireDPPS.Enabled = True

End Sub

Private Sub cmdLoad_Click()

If (CheckROTATE.value = vbChecked) Then
        MARK_ANGLE = MARK_ANGLE_ROTATED
Else
        MARK_ANGLE = MARK_ANGLE_DEFAULT
End If
 
Load_Job_From_File
    
End Sub

Private Sub cmdM3_Click()

txtFrequency.Text = Val(txtFrequency.Text) - 0.01
If Val(txtFrequency.Text) < 0.02 Then
    txtFrequency.Text = 0.02
End If

End Sub

Private Sub cmdMinus1_Click()
txtZHT.Text = txtZHT.Text + 1
End Sub

Private Sub cmdMinus11_Click()
txtZHT.Text = txtZHT.Text + 10
End Sub

Private Sub cmdP3_Click()

txtFrequency.Text = Val(txtFrequency.Text) + 0.01
If Val(txtFrequency.Text) > 250 Then
    txtFrequency.Text = 250
End If

End Sub

Private Sub cmdPlus1_Click()
txtZHT.Text = txtZHT.Text - 1
End Sub

Private Sub cmdPlus11_Click()
txtZHT.Text = txtZHT.Text - 10
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

Private Sub cmdRefresh_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [SIZE_LOC_ID],[CASE SIZE],mid([CASE NAME],1,5),format([TEXT XSIZE],'0.000'),format([TEXT YSIZE],'0.000') From [TBL SIZE LOC]"

sSQLF = "       |^SIZE_ID |^Case|<                |^Horiz       |^Vert     "

Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

End Sub


Private Sub cmdSetObjCharString_Click()

'=============================================
'   LOGO MODE
'=============================================

If (optLogo1.value = True) Then
    LOGO_MODE = LOGO_SIDE
End If
If (optLogo2.value = True) Then
    LOGO_MODE = LOGO_TOP
End If
If (optLogo3.value = True) Then
    LOGO_MODE = LOGO_OMIT
End If
If (optLogo4.value = True) Then
    LOGO_MODE = LOGO_ATC
End If

'=============================================
'   MARK_MODE
'=============================================

If (OptLine(0).value = True) Then
     MARK_MODE = 1
End If
If (OptLine(1).value = True) Then
      MARK_MODE = 2
End If
If (OptLine(2).value = True) Then
      MARK_MODE = 3
End If
If (OptLine(3).value = True) Then
      MARK_MODE = 4
End If

Dim sBuff As String
'=============================================
'[1] SetObjCharString
'=============================================

Select Case MARK_MODE
Case 1
        sBuff = Text1.Text
        AutomationInterface.SetObjCharString 0, sBuff
Case 2
        sBuff = Text1.Text
        AutomationInterface.SetObjCharString 0, sBuff
        sBuff = Text2.Text
        AutomationInterface.SetObjCharString 1, sBuff
Case 3
        sBuff = Text1.Text
        AutomationInterface.SetObjCharString 0, sBuff
        sBuff = Text2.Text
        AutomationInterface.SetObjCharString 1, sBuff
        sBuff = Text3.Text
        AutomationInterface.SetObjCharString 2, sBuff
Case 4
        sBuff = Text1.Text
        AutomationInterface.SetObjCharString 0, sBuff
        sBuff = Text2.Text
        AutomationInterface.SetObjCharString 1, sBuff
        sBuff = Text3.Text
        AutomationInterface.SetObjCharString 2, sBuff
        sBuff = Text4.Text
        AutomationInterface.SetObjCharString 3, sBuff
End Select

cmdSetObjCharString.FontBold = True

'MsgBox "SetObjCharString Complete", vbInformation, "Laser"

End Sub

Private Sub cmdSetObjProfile_Click()

Data2.UpdateRecord

Dim sSQL As String
Set FR_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT * FROM [TBL POWER] WHERE [TBL_ID] =" & POWER_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

ProfileIndex = 0

Markspeed = FR_Table.Fields("[Markspeed]")
Jumpspeed = FR_Table.Fields("[Jumpspeed]")

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

If (optLogo5.value = True) Then
    LOGO_MODE = LOGO_ABRASIVE
    Markspeed = FR_Table.Fields("[ABRASIZE_Markspeed]")
    Markspeed_Bits = Format(Markspeed * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")
    
    T1 = FR_Table.Fields("[ABRASIZE_Frequency]")
    T2 = FR_Table.Fields("[ABRASIZE_PulseWidth]")
End If

Dim i As Integer
For i = 0 To ObjectCount - 1
        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        AutomationInterface.SetObjProfile i, ProfileIndex, Markspeed_Bits, Jumpspeed_Bits, Jumpdelay, Markdelay, Polygondelay, Laserpower, Laseroffdelay, Laserondelay, Eightbitword, T1, T2, zAxis, Varijumpdelay, Varijumplength, Wobblesize, Wobblefrequency, Powerreset, Varipolydelay
Next i

sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)
  
ObjHPosition(0) = FR_Table.Fields("[TXT1 X]")
ObjVPosition(0) = FR_Table.Fields("[TXT1 Y]")
ObjHPosition(1) = FR_Table.Fields("[TXT2 X]")
ObjVPosition(1) = FR_Table.Fields("[TXT2 Y]")
ObjHPosition(2) = FR_Table.Fields("[TXT3 X]")
ObjVPosition(2) = FR_Table.Fields("[TXT3 Y]")
ObjHPosition(3) = FR_Table.Fields("[TXT4 X]")
ObjVPosition(3) = FR_Table.Fields("[TXT4 Y]")
   
Select Case LOGO_MODE
Case LOGO_SIDE
    ObjHPosition(4) = FR_Table.Fields("[LOGO LX SIDE]")
    ObjVPosition(4) = FR_Table.Fields("[LOGO LY SIDE]")
Case LOGO_TOP
    ObjHPosition(4) = FR_Table.Fields("[LOGO LX TOP]")
    ObjVPosition(4) = FR_Table.Fields("[LOGO LY TOP]")
Case LOGO_ATC
    ObjHPosition(5) = FR_Table.Fields("[ATC  X]")
    ObjVPosition(5) = FR_Table.Fields("[ATC  Y]")
Case LOGO_ABRASIVE
    'ABRASIZE
    ObjHPosition(6) = FR_Table.Fields("[LINE X1]")
    ObjVPosition(6) = FR_Table.Fields("[LINE Y1]")
    ObjHPosition(7) = FR_Table.Fields("[LINE X2]")
    ObjVPosition(7) = FR_Table.Fields("[LINE Y2]")
        
    ObjHPosition(8) = FR_Table.Fields("[LINE X1]") + (FR_Table.Fields("[LEN H]") * 0.5) - (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(8) = FR_Table.Fields("[LINE Y1]")
    ObjHPosition(9) = FR_Table.Fields("[LINE X1]") + (FR_Table.Fields("[LEN H]") * 0.5) + (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(9) = FR_Table.Fields("[LINE Y1]")
    
    ObjHPosition(10) = FR_Table.Fields("[LINE X2]") + (FR_Table.Fields("[LEN H]") * 0.5) - (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(10) = FR_Table.Fields("[LINE Y2]") - FR_Table.Fields("[LEN V]")
    ObjHPosition(11) = FR_Table.Fields("[LINE X2]") + (FR_Table.Fields("[LEN H]") * 0.5) + (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(11) = FR_Table.Fields("[LINE Y2]") - FR_Table.Fields("[LEN V]")
    
    REP_ABRASIVE = FR_Table.Fields("[REP]")
End Select

Select Case MARK_ANGLE
Case MARK_ANGLE_DEFAULT
        
Case MARK_ANGLE_ROTATED
            ObjHPosition(0) = FR_Table.Fields("[R TXT1 X]")
            ObjVPosition(0) = FR_Table.Fields("[R TXT1 Y]")
            ObjHPosition(1) = FR_Table.Fields("[R TXT2 X]")
            ObjVPosition(1) = FR_Table.Fields("[R TXT2 Y]")
            ObjHPosition(2) = FR_Table.Fields("[R TXT3 X]")
            ObjVPosition(2) = FR_Table.Fields("[R TXT3 Y]")
            ObjHPosition(3) = FR_Table.Fields("[R TXT4 X]")
            ObjVPosition(3) = FR_Table.Fields("[R TXT4 Y]")
            Select Case LOGO_MODE
            Case LOGO_SIDE
                ObjHPosition(4) = FR_Table.Fields("[R LOGO LX SIDE]")
                ObjVPosition(4) = FR_Table.Fields("[R LOGO LY SIDE]")
            Case LOGO_TOP
                ObjHPosition(4) = FR_Table.Fields("[R LOGO LX TOP]")
                ObjVPosition(4) = FR_Table.Fields("[R LOGO LY TOP]")
            Case LOGO_ATC
                ObjHPosition(5) = FR_Table.Fields("[R ATC  X]")
                ObjVPosition(5) = FR_Table.Fields("[R ATC  Y]")
            Case LOGO_ABRASIVE

            End Select
End Select

FR_Table.Close
FR_Database.Close

'MsgBox "SetObjProfiles Complete", vbInformation, "Laser"
cmdSetObjProfile.FontBold = True
End Sub

Private Sub cmdSetObjSize_Click()

Data3.UpdateRecord
'=============================================
'   LOGO MODE
'=============================================

If (optLogo2.value = True) Then
    LOGO_MODE = LOGO_ATC
End If
If (optLogo3.value = True) Then
    LOGO_MODE = LOGO_TOP
End If
If (optLogo1.value = True) Then
    LOGO_MODE = LOGO_SIDE
End If
If (optLogo4.value = True) Then
    LOGO_MODE = LOGO_OMIT
End If
If (optLogo5.value = True) Then
    LOGO_MODE = LOGO_ABRASIVE
End If

'=============================================
'   MARK_MODE
'=============================================

If (OptLine(0).value = True) Then
     MARK_MODE = 1
End If
If (OptLine(1).value = True) Then
      MARK_MODE = 2
End If
If (OptLine(2).value = True) Then
      MARK_MODE = 3
End If
If (OptLine(3).value = True) Then
      MARK_MODE = 4
End If

'=============================================
'[2] SetObjSize
'=============================================

Dim sSQL As String
Set FR_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)


Dim HTextSize As Long
Dim VTextSize As Long

HTextSize = Format(FR_Table.Fields("[TEXT XSIZE]") * INCHES_TO_BITS, "0")
VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS, "0")

AutomationInterface.SetObjSize 0, HTextSize, VTextSize
AutomationInterface.SetObjSize 1, HTextSize, VTextSize
AutomationInterface.SetObjSize 2, HTextSize, VTextSize
AutomationInterface.SetObjSize 3, HTextSize, VTextSize

Select Case LOGO_MODE
Case LOGO_SIDE, LOGO_TOP
                HTextSize = Format(FR_Table.Fields("[LOGO XSIZE]") * INCHES_TO_BITS, "0")
                VTextSize = Format(FR_Table.Fields("[LOGO YSIZE]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 4, HTextSize, VTextSize
Case LOGO_OMIT

Case LOGO_ATC
                HTextSize = Format(FR_Table.Fields("[TEXT XSIZE]") * INCHES_TO_BITS, "0")
                VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 5, HTextSize, VTextSize
                
Case LOGO_ABRASIVE
                'SIZE = LENGTH
                HTextSize = Format(FR_Table.Fields("[LEN H]") * INCHES_TO_BITS, "0")
                VTextSize = 0
                AutomationInterface.SetObjSize 6, HTextSize, VTextSize
                AutomationInterface.SetObjSize 7, HTextSize, VTextSize

                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                HTextSize = 0
                AutomationInterface.SetObjSize 8, HTextSize, VTextSize

                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 9, HTextSize, VTextSize

                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 10, HTextSize, VTextSize
                                
                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 11, HTextSize, VTextSize
End Select

FR_Table.Close
FR_Database.Close

cmdSetObjSize.FontBold = True
'MsgBox "SetObjSize Complete", vbInformation, "Laser"

End Sub

Private Sub cmdSetup_Click()

'=============================================
'   LOGO MODE
'=============================================

If (optLogo1.value = True) Then
    LOGO_MODE = LOGO_SIDE
End If
If (optLogo2.value = True) Then
    LOGO_MODE = LOGO_TOP
End If
If (optLogo3.value = True) Then
    LOGO_MODE = LOGO_OMIT
End If
If (optLogo4.value = True) Then
    LOGO_MODE = LOGO_ATC
End If

'=============================================
'   MARK_MODE
'=============================================

If (OptLine(0).value = True) Then
     MARK_MODE = 1
End If
If (OptLine(1).value = True) Then
      MARK_MODE = 2
End If
If (OptLine(2).value = True) Then
      MARK_MODE = 3
End If
If (OptLine(3).value = True) Then
      MARK_MODE = 4
End If

Dim sBuff As String
'=============================================
'[1] SETUP TEXT CHAR STRING SetObjCharString
'=============================================
Select Case MARK_MODE
Case 1, 2, 3, 4
        sBuff = Text4.Text
        AutomationInterface.SetObjCharString 3, sBuff
End Select
Select Case MARK_MODE
Case 1, 2, 3
        sBuff = Text3.Text
        AutomationInterface.SetObjCharString 2, sBuff
End Select
Select Case MARK_MODE
Case 1, 2
        sBuff = Text2.Text
        AutomationInterface.SetObjCharString 1, sBuff
End Select
Select Case MARK_MODE
Case 1
        sBuff = Text1.Text
        AutomationInterface.SetObjCharString 0, sBuff
End Select

'=============================================
'[2] SETUP TEXT SIZE SetObjSize
'=============================================

Dim sSQL As String
Set FR_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)


Dim HTextSize As Long
Dim VTextSize As Long

HTextSize = Format(FR_Table.Fields("[TEXT XSIZE]") * INCHES_TO_BITS, "0")
VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS, "0")

AutomationInterface.SetObjSize 0, HTextSize, VTextSize
AutomationInterface.SetObjSize 1, HTextSize, VTextSize
AutomationInterface.SetObjSize 2, HTextSize, VTextSize
AutomationInterface.SetObjSize 3, HTextSize, VTextSize

Select Case LOGO_MODE
Case LOGO_SIDE, LOGO_TOP
                HTextSize = Format(FR_Table.Fields("[LOGO XSIZE]") * INCHES_TO_BITS, "0")
                VTextSize = Format(FR_Table.Fields("[LOGO YSIZE]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 4, HTextSize, VTextSize
Case LOGO_OMIT

Case LOGO_ATC
                HTextSize = Format(FR_Table.Fields("[TEXT XSIZE]") * INCHES_TO_BITS, "0")
                VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 5, HTextSize, VTextSize
End Select

FR_Table.Close
FR_Database.Close

End Sub

Private Sub cmdUpdate_Click()
Data2.UpdateRecord
CommandRefresh_Click
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

Private Sub cmdUpdate2_Click()
Data3.UpdateRecord
End Sub

Private Sub cmdUpdateRecord2_Click()

End Sub

Private Sub cmdZHT_Click()
MoveToTarget 3, CLng(txtZHT.Text)
End Sub

Private Sub Command1_Click()

Dim Xctr As Long
Dim Yctr As Long

Xctr = Format(Val(TextX.Text) * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")
Yctr = Format(Val(TextY.Text) * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")

Dim i As Integer
For i = 0 To ObjectCount - 1
        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        AutomationInterface.RotateObjEx i, Val(TextAngle.Text), Xctr, Yctr
Next i

MsgBox "RotateObjEx Complete", vbInformation, "Laser"

End Sub

Private Sub Command2_Click()
FindReverseLimit 3
End Sub

Private Sub CommandRefresh_Click()

MSFlexGrid3.Height = 1000

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TBL_ID],[SERIES],[CASE],[VALUE],[COATING],[ATC PART]," & _
              "format([Frequency],'0.00')," & _
              "format([Markspeed],'0.000')," & _
              "format([PulseWidth],'0.00')," & _
                     "[ANGLE],[Z HEIGHT] " & _
       "FROM [TBL Power] " & _
       "WHERE [ACTIVE] = Yes AND " & _
             "[CASE]    ='" & CASE_ID & "' AND " & _
             "[TRAY_ID]  =" & TRAY_ID
                                   
sSQLF = "   |^ID  |<Series       |^Case|^Range         |<Coating          |<ATC Part        |>Frequency|>Markspeed|>PulseWidth|^Angle    |<ZHT   "


sSQL = "SELECT [TBL_ID],[SERIES],[CASE],[VALUE],[COATING],[ATC PART]," & _
              "format([Frequency],'0.00')," & _
              "format([Markspeed],'0.000')," & _
              "format([PulseWidth],'0.00')," & _
                     "[ANGLE],[Z HEIGHT] " & _
       "FROM [TBL POWER] WHERE [TBL_ID] =" & POWER_ID
                                   
sSQLF = "   |^POWER_ID|<Series       |^Case|^Range         |<Coating          |<ATC Part        |>Frequency|>Markspeed|>PulseWidth|^Angle    |<ZHT   "


Data4.RecordSource = sSQL
Data4.Refresh
 
MSFlexGrid3.FormatString = sSQLF

End Sub

Private Sub CommandRotateObj_Click()

Dim i As Integer
For i = 0 To ObjectCount - 1
        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        AutomationInterface.RotateObj i, Val(TextAngle.Text)
Next i

MsgBox "RotateObj Complete", vbInformation, "Laser"

End Sub

'
'
Private Sub Form_Load()

Caption = "Power & Scale Factors           " & ATC_DWG & "         " & ATC_VERSION

'If (Load_Job = 1) Then
'    cmdLoad.Visible = False
'End If

Data1.DatabaseName = ATC_LASER_BD
Data2.DatabaseName = ATC_LASER_BD
Data3.DatabaseName = ATC_LASER_BD
Data4.DatabaseName = ATC_LASER_BD

Dim sSQL As String

sSQL = "SELECT * FROM [TBL POWER] WHERE [TBL_ID] =" & POWER_ID

Data2.RecordSource = sSQL
Data2.Refresh
                                   
sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID
                                   
Data3.RecordSource = sSQL
Data3.Refresh

txtStartNo.Text = SERIAL_START_NUMBER

LabelSIZE_LOC_ID.Caption = SIZE_LOC_ID
LabelTRAY_ID.Caption = TRAY_ID
LabelPOWER_ID.Caption = POWER_ID

Select Case TRAY_MARK_ANGLE
Case MARK_ANGLE_DEFAULT
        LabelMARK_ANGLE.Caption = "Default"
Case MARK_ANGLE_ROTATED
        LabelMARK_ANGLE.Caption = "Rotated"
End Select

cmdRefresh_Click

CommandRefresh_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

Select Case OP_MODE
Case 0
        frmMain.Show
Case 1
        frmOPScreen.Show
End Select

End Sub

Private Sub txtScale_GotFocus(Index As Integer)

txtScale(Index).SelStart = 0
txtScale(Index).SelLength = Len(txtScale(Index))

End Sub
