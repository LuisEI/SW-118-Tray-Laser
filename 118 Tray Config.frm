VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTray 
   Caption         =   "118 Tray Configuration"
   ClientHeight    =   11880
   ClientLeft      =   225
   ClientTop       =   795
   ClientWidth     =   20955
   LinkTopic       =   "Form1"
   ScaleHeight     =   11880
   ScaleWidth      =   20955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text22 
      BackColor       =   &H00FFFFC0&
      DataField       =   "CAMERA POS"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   4800
      TabIndex        =   150
      Text            =   "1"
      ToolTipText     =   "CAMERA POS"
      Top             =   11640
      Width           =   1335
   End
   Begin VB.CommandButton CommandUpdateHeight 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Global Height"
      Height          =   300
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   11160
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   11880
      TabIndex        =   136
      Top             =   3960
      Width           =   5295
      Begin VB.TextBox txtSteps 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Steps"
         DataSource      =   "Data3"
         Height          =   300
         Left            =   360
         TabIndex        =   143
         Text            =   "30.0"
         ToolTipText     =   "Steps"
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox txtInches 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Inches"
         DataSource      =   "Data3"
         Height          =   300
         Left            =   1200
         TabIndex        =   142
         Text            =   "30.0"
         ToolTipText     =   "Inches"
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox Text1 
         DataField       =   "RANGE MIN"
         DataSource      =   "Data3"
         Height          =   300
         Left            =   2040
         TabIndex        =   141
         Text            =   "30.0"
         ToolTipText     =   "RANGE MIN"
         Top             =   1920
         Width           =   435
      End
      Begin VB.TextBox Text3 
         DataField       =   "RANGE MAX"
         DataSource      =   "Data3"
         Height          =   300
         Left            =   2760
         TabIndex        =   140
         Text            =   "30.0"
         ToolTipText     =   "RANGE MAX"
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox Text4 
         DataField       =   "HOME"
         DataSource      =   "Data3"
         Height          =   300
         Left            =   3600
         TabIndex        =   139
         Text            =   "30.0"
         ToolTipText     =   "HOME"
         Top             =   1920
         Width           =   315
      End
      Begin VB.CommandButton CommandUpdateRecord 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Update Record"
         Height          =   300
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   2400
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   855
         Left            =   240
         TabIndex        =   137
         ToolTipText     =   "FROM  [TBL Axis]"
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1508
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin VB.Label Label4 
         Caption         =   "Steps"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   360
         TabIndex        =   148
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Inches"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1200
         TabIndex        =   147
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "MIN"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   146
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "MAX"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2760
         TabIndex        =   145
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Load/Home"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3600
         TabIndex        =   144
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5] Order"
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
      Left            =   480
      TabIndex        =   135
      Top             =   11520
      Width           =   2000
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H00FFFFC0&
      DataField       =   "ORDER"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   11880
      TabIndex        =   133
      Text            =   "1"
      ToolTipText     =   "CASE"
      Top             =   10200
      Width           =   375
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 11"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   11040
      TabIndex        =   131
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   11640
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 10 NON"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   11040
      TabIndex        =   129
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   11280
      Width           =   1935
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 10 MAG"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   11040
      TabIndex        =   127
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   10920
      Width           =   1935
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 9"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   11040
      TabIndex        =   125
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   10560
      Width           =   1935
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H00FFFFC0&
      DataField       =   "STAR"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   10200
      TabIndex        =   123
      Text            =   "1"
      ToolTipText     =   "CASE"
      Top             =   10200
      Width           =   375
   End
   Begin VB.CommandButton CommandExcel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Excel"
      Height          =   300
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   10680
      Width           =   1815
   End
   Begin VB.TextBox txtLocation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      DataField       =   "CASE"
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
      Index           =   23
      Left            =   4080
      MaxLength       =   1
      TabIndex        =   120
      Text            =   "1"
      ToolTipText     =   "CASE"
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      DataField       =   "ATC DWG"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   118
      Text            =   "Title"
      ToolTipText     =   "ATC DWG"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtLocation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      DataField       =   "ROTATION"
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
      Index           =   22
      Left            =   8880
      MaxLength       =   1
      TabIndex        =   116
      Text            =   "1"
      ToolTipText     =   "ROTATION"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton CommandInitialize 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Initialize"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFC0FF&
      Caption         =   ">>>"
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
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFC0FF&
      Caption         =   "<<<"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFC0FF&
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
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFC0FF&
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
      Left            =   6765
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFC0FF&
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
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFC0FF&
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
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   5280
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laser Offset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      TabIndex        =   98
      Top             =   5280
      Width           =   2655
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFC0FF&
         DataField       =   "L Y Offset"
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
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   102
         Text            =   "30.0"
         ToolTipText     =   "L Y Offset"
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFC0FF&
         DataField       =   "L X Offset"
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
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   101
         Text            =   "30.0"
         ToolTipText     =   "L X Offset"
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFC0FF&
         DataField       =   "L X Offset 1"
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
         Height          =   375
         Index           =   9
         Left            =   600
         TabIndex        =   100
         Text            =   "30.0"
         ToolTipText     =   "L X Offset 1"
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFC0FF&
         DataField       =   "L Y Offset 1"
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
         Height          =   375
         Index           =   10
         Left            =   1800
         TabIndex        =   99
         Text            =   "30.0"
         ToolTipText     =   "L Y Offset 1"
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label49 
         Caption         =   "Y1"
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
         Left            =   1440
         TabIndex        =   106
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label48 
         Caption         =   "X1"
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
         TabIndex        =   105
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "X0"
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
         TabIndex        =   104
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Y0"
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
         Left            =   1440
         TabIndex        =   103
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton CommandAbort 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stop Motion"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox TextSegment 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Left            =   2040
      TabIndex        =   95
      Text            =   "1"
      ToolTipText     =   "L X Offset"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rev Limit Tray"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   5820
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Goto Segment"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Frame FramePWR 
      Caption         =   " Laser Parameters "
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
      Height          =   2895
      Left            =   14280
      TabIndex        =   80
      Top             =   7200
      Width           =   2655
      Begin VB.TextBox txtFrequency 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Frequency"
         DataSource      =   "Data7"
         Height          =   285
         Left            =   1680
         TabIndex        =   84
         Text            =   "XXX"
         Top             =   645
         Width           =   735
      End
      Begin VB.TextBox txtMarkspeed 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Markspeed"
         DataSource      =   "Data7"
         Height          =   285
         Left            =   1680
         TabIndex        =   83
         Text            =   "XXXXXX"
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtPulseWidth 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PulseWidth"
         DataSource      =   "Data7"
         Height          =   285
         Left            =   1680
         TabIndex        =   82
         Text            =   "X"
         ToolTipText     =   "PulseWidth"
         Top             =   1785
         Width           =   735
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ANGLE"
         DataSource      =   "Data7"
         Height          =   285
         Left            =   1680
         TabIndex        =   81
         Text            =   "X"
         ToolTipText     =   "ANGLE"
         Top             =   2355
         Width           =   735
      End
      Begin VB.Label Label38 
         Caption         =   "Frequency   (kHz)"
         Height          =   285
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "[0.02 to 250.0]"
         Height          =   285
         Left            =   1320
         TabIndex        =   91
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Mark Speed  (in/ms)"
         Height          =   285
         Left            =   120
         TabIndex        =   90
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label41 
         Caption         =   "Pulse Width  (us)"
         Height          =   285
         Left            =   120
         TabIndex        =   89
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "[0 to 30000]"
         Height          =   285
         Left            =   1560
         TabIndex        =   88
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "[2 to 65535] "
         Height          =   285
         Left            =   1440
         TabIndex        =   87
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label44 
         Caption         =   "Angle  [degrees]"
         Height          =   285
         Left            =   120
         TabIndex        =   86
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "[-360.00 to 360.00] "
         Height          =   285
         Left            =   840
         TabIndex        =   85
         Top             =   2070
         Width           =   1575
      End
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4] ID Active"
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
      Left            =   480
      TabIndex        =   79
      Top             =   11160
      Width           =   2000
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL10"
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
      Index           =   21
      Left            =   10920
      TabIndex        =   77
      Text            =   "10"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL9"
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
      Index           =   20
      Left            =   10200
      TabIndex        =   75
      Text            =   "9"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL8"
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
      Index           =   19
      Left            =   9480
      TabIndex        =   73
      Text            =   "8"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL7"
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
      Index           =   18
      Left            =   8760
      TabIndex        =   71
      Text            =   "7"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL6"
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
      Index           =   17
      Left            =   8040
      TabIndex        =   69
      Text            =   "6"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL5"
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
      Index           =   16
      Left            =   7320
      TabIndex        =   67
      Text            =   "5"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL4"
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
      Index           =   15
      Left            =   6600
      TabIndex        =   65
      Text            =   "4"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL3"
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
      Index           =   14
      Left            =   5880
      TabIndex        =   63
      Text            =   "3"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL2"
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
      Index           =   13
      Left            =   5160
      TabIndex        =   61
      Text            =   "2"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "COL1"
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
      Index           =   12
      Left            =   4440
      TabIndex        =   59
      Text            =   "1"
      ToolTipText     =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtLocation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      DataField       =   "PAGE"
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
      Index           =   11
      Left            =   6120
      MaxLength       =   1
      TabIndex        =   57
      Text            =   "1"
      ToolTipText     =   "PAGE"
      Top             =   4560
      Width           =   375
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3]  TRAY_ID"
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
      Left            =   480
      TabIndex        =   56
      Top             =   10800
      Value           =   -1  'True
      Width           =   2000
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Active"
      DataField       =   "ACTIVE"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   3120
      TabIndex        =   55
      Top             =   11340
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Record"
      Height          =   300
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   10320
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFC0&
      DataField       =   "DV MAX"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   7680
      TabIndex        =   52
      Text            =   "1"
      ToolTipText     =   "DV MAX"
      Top             =   11370
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFC0&
      DataField       =   "DV MIN"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   7680
      TabIndex        =   50
      Text            =   "1"
      ToolTipText     =   "DV MIN"
      Top             =   11085
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFC0&
      DataField       =   "VALUE"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   7680
      TabIndex        =   48
      Text            =   "1"
      ToolTipText     =   "Spacing INDEX"
      Top             =   10800
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFC0&
      DataField       =   "CASE "
      DataSource      =   "Data7"
      Height          =   285
      Left            =   7680
      TabIndex        =   46
      Text            =   "1"
      ToolTipText     =   "CASE"
      Top             =   10515
      Width           =   375
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFC0&
      DataField       =   "SERIES"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   7680
      TabIndex        =   44
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   10230
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFC0&
      DataField       =   "COATING"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   4800
      TabIndex        =   42
      Text            =   "1"
      ToolTipText     =   "COATING"
      Top             =   11055
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFC0&
      DataField       =   "ATC PART"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   4800
      TabIndex        =   40
      Text            =   "1"
      ToolTipText     =   "ATC PART"
      Top             =   10770
      Width           =   1335
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "SEG Y DIST"
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
      Index           =   8
      Left            =   10560
      TabIndex        =   38
      Text            =   "30.0"
      Top             =   6600
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] TRAY_ID All"
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
      Left            =   480
      TabIndex        =   37
      Top             =   10080
      Width           =   2000
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2]  DWG"
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
      Left            =   480
      TabIndex        =   36
      Top             =   10440
      Width           =   2000
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      DataField       =   "ATC DWG"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   4800
      TabIndex        =   34
      Text            =   "1"
      ToolTipText     =   "ATC DWG"
      Top             =   10485
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      DataField       =   "TRAY_ID"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   32
      Text            =   "12"
      ToolTipText     =   "Spacing INDEX"
      Top             =   10200
      Width           =   615
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL POWER"
      Top             =   8520
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   8040
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.CommandButton cmdRefresh6 
      Caption         =   "Refresh6"
      Height          =   300
      Left            =   5160
      TabIndex        =   30
      Top             =   9240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      DataField       =   "TITLE"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   29
      Text            =   "Title"
      ToolTipText     =   "TITLE"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "X 0"
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
      Index           =   0
      Left            =   6720
      TabIndex        =   27
      Text            =   "30.0"
      ToolTipText     =   "X 0"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Stop"
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
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh4 
      Caption         =   "Refresh4"
      Height          =   250
      Left            =   14040
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data8 
      Caption         =   "Data8 [TBL Axis]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Axis"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Move"
      Height          =   300
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   1215
   End
   Begin VB.VScrollBar vsbCP 
      Height          =   1815
      Left            =   16440
      Max             =   1000
      TabIndex        =   21
      Top             =   840
      Value           =   1
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reverse Limit"
      Height          =   300
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "flex_find_reference"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Move to Camera Offset"
      Height          =   300
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtCameraPosition 
      BackColor       =   &H00C0FFC0&
      DataField       =   "Y Offset A"
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
      Height          =   375
      Left            =   15240
      TabIndex        =   17
      Text            =   "30.0"
      ToolTipText     =   "Camera Position"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 [TBL Axis]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Axis"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.TextBox txtSpacing 
      BackColor       =   &H00FFC0FF&
      DataField       =   "Spacing"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      Text            =   "30.0"
      ToolTipText     =   "Spacing"
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Calculate"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6660
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh1 
      Caption         =   "Refresh1"
      Height          =   300
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Update Record"
      Height          =   300
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "X 5"
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
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Text            =   "30.0"
      ToolTipText     =   "X 5"
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "X 4"
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
      Index           =   4
      Left            =   3960
      TabIndex        =   3
      Text            =   "30.0"
      ToolTipText     =   "X 4"
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "X 3"
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
      Index           =   3
      Left            =   3960
      TabIndex        =   2
      Text            =   "30.0"
      ToolTipText     =   "X 3"
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "X 1"
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
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Text            =   "30.0"
      ToolTipText     =   "X 1"
      Top             =   5400
      Width           =   975
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 [TBL Tray Configs]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\118 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   2280
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [TBL Tray Config]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFC0FF&
      DataField       =   "X 2"
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
      Index           =   2
      Left            =   3960
      TabIndex        =   0
      Text            =   "30.0"
      ToolTipText     =   "X 2"
      Top             =   5760
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "FROM  [TBL Tray Config]"
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1508
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
      Height          =   2655
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "FROM [TBL Power]"
      Top             =   7320
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   4683
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      Caption         =   "FROM [TBL Power]"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   13680
      TabIndex        =   152
      Top             =   11760
      Width           =   1395
   End
   Begin VB.Label Label60 
      Caption         =   "CAMERA POS"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   151
      Top             =   11760
      Width           =   1335
   End
   Begin VB.Label Label59 
      Caption         =   "Order"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11040
      TabIndex        =   134
      Top             =   10200
      Width           =   600
   End
   Begin VB.Label Label58 
      Caption         =   "POS 11"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9360
      TabIndex        =   132
      Top             =   11640
      Width           =   1200
   End
   Begin VB.Label Label57 
      Caption         =   "POS 10 NON MAG "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9360
      TabIndex        =   130
      Top             =   11280
      Width           =   1560
   End
   Begin VB.Label Label56 
      Caption         =   "POS 10 MAG"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9360
      TabIndex        =   128
      Top             =   10920
      Width           =   1200
   End
   Begin VB.Label Label55 
      Caption         =   "POS 9"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9360
      TabIndex        =   126
      Top             =   10560
      Width           =   1200
   End
   Begin VB.Label Label54 
      Caption         =   "STAR"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9360
      TabIndex        =   124
      Top             =   10200
      Width           =   600
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      Caption         =   "Case"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3240
      TabIndex        =   121
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      Caption         =   "ATC DWG : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   119
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      Caption         =   "Rotation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7800
      TabIndex        =   117
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8280
      TabIndex        =   114
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   113
      Top             =   5280
      Width           =   135
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Caption         =   "Tray Title : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   96
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label37 
      Caption         =   "Col 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   10920
      TabIndex        =   78
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label36 
      Caption         =   "Col 9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   10200
      TabIndex        =   76
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label35 
      Caption         =   "Col 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   9480
      TabIndex        =   74
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label34 
      Caption         =   "Col 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   8760
      TabIndex        =   72
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label33 
      Caption         =   "Col 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   8040
      TabIndex        =   70
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label32 
      Caption         =   "Col 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7320
      TabIndex        =   68
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label31 
      Caption         =   "Col 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6600
      TabIndex        =   66
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label30 
      Caption         =   "Col 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5880
      TabIndex        =   64
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label29 
      Caption         =   "Col 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      TabIndex        =   62
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "Col 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   4440
      TabIndex        =   60
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5280
      TabIndex        =   58
      Top             =   4560
      Width           =   735
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   11145
      Left            =   18240
      Top             =   120
      Width           =   2130
   End
   Begin VB.Label Label27 
      Caption         =   "DV MAX"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6480
      TabIndex        =   53
      Top             =   11370
      Width           =   1200
   End
   Begin VB.Label Label26 
      Caption         =   "DV MIN"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6480
      TabIndex        =   51
      Top             =   11085
      Width           =   1200
   End
   Begin VB.Label Label25 
      Caption         =   "Value Range"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6480
      TabIndex        =   49
      Top             =   10800
      Width           =   1200
   End
   Begin VB.Label Label24 
      Caption         =   "CASE"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6480
      TabIndex        =   47
      Top             =   10515
      Width           =   1200
   End
   Begin VB.Label Label23 
      Caption         =   "SERIES"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6480
      TabIndex        =   45
      Top             =   10230
      Width           =   1200
   End
   Begin VB.Label Label22 
      Caption         =   "COATING"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   43
      Top             =   11055
      Width           =   1335
   End
   Begin VB.Label Label21 
      Caption         =   "ATC PART"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   41
      Top             =   10770
      Width           =   1335
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Seg Y Distance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   8760
      TabIndex        =   39
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "ATC Dwg"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   35
      Top             =   10485
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "TRAY_ID [1-12]"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   33
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "Load [X0]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5520
      TabIndex        =   28
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "X 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3000
      TabIndex        =   26
      Top             =   6840
      Width           =   795
   End
   Begin VB.Label lblCP 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   300
      Left            =   15480
      TabIndex        =   22
      ToolTipText     =   "Camera Position"
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Axis 2 Camera Alignment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   14040
      TabIndex        =   20
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Segments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Spacing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "X 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3000
      TabIndex        =   14
      Top             =   6480
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "X 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3000
      TabIndex        =   13
      Top             =   6120
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "X 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3000
      TabIndex        =   10
      Top             =   5760
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "X 1 [Master]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2640
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblSegments 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      DataField       =   "Segments"
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Segments"
      Top             =   4080
      Width           =   375
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdCalculate_Click()

txtLocation(2).Text = Format(Val(txtLocation(1).Text) - Val(txtSpacing.Text) * Val(txtSteps.Text) / Val(txtInches.Text), "0")
txtLocation(3).Text = Format(Val(txtLocation(1).Text) - 2 * Val(txtSpacing.Text) * Val(txtSteps.Text) / Val(txtInches.Text), "0")
txtLocation(4).Text = Format(Val(txtLocation(1).Text) - 3 * Val(txtSpacing.Text) * Val(txtSteps.Text) / Val(txtInches.Text), "0")

End Sub

Private Sub cmdMove_Click()

MoveToTarget 2, CLng(txtCameraPosition.Text)

End Sub

Private Sub cmdRefresh1_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TRAY_ID],[CASE],[TITLE],[ROWS]& ' X ' & [COLS],[ATC DWG]," & _
                "[Segments],[Spacing]," & _
                "[L X OffSet],[L Y OffSet]," & _
                "[X 0],[X 1],[X 2],[X 3],[X 4],[X 5],[PAGE],[Y OffSet A],[ROTATION] " & _
        "FROM  [TBL Tray Config]"
 

sSQLF = "   |^T_ID|^Case|<Tray Title                          |^Row X Col |^ATC DWG  |Segments|>Spacing|>X Offset|>Y Offset"
sSQLF = sSQLF & "|^Load     |^X1       |^X2       |^X3       |^X4       |^X5        |^Page|>CP       |^Rot"

Data1.RecordSource = sSQL
Data1.Refresh
 
MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdRefresh4_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [AXIS_ID],[DESC],[Steps],[Inches]," & _
                "[RANGE MIN],[RANGE MAX],[HOME] " & _
        "FROM  [TBL Axis] "
                    
sSQLF = "   |^AXIS|<Description|STEPS|^INCHES|>MIN|>MAX     |>HOME  "

Data8.RecordSource = sSQL
Data8.Refresh
 
MSFlexGrid4.FormatString = sSQLF

End Sub

Private Sub cmdRefresh6_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TRAY_ID],[TBL_ID],[ATC DWG]," & _
               "format([Frequency],'0.0')," & _
               "format([Markspeed],'0.00')," & _
               "format([PulseWidth],'0.00'),[ANGLE]," & _
"[TITLE],[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR],[Z HEIGHT],[HEIGHT],format([ACTIVE],'Yes/No'),[CAMERA POS],[MARK PARA]" & _
       "FROM [TBL Power] WHERE [TRAY_ID]=" & TRAY_ID

If (Option1.value = True) Then
    sSQL = sSQL & " ORDER BY [TRAY_ID],[ATC PART],[COATING],[SERIES]"
End If
If (Option2.value = True) Then
    sSQL = sSQL & "AND [ACTIVE] = Yes  ORDER BY [ATC DWG]"
End If
If (Option3.value = True) Then
    sSQL = sSQL & "AND [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ATC PART],[COATING]"
End If

sSQLF = "   |^TRAY_ID|^PID|^ATC DWG|>Frequency|>Markspeed|>PulseWidth|^Angle    |<Title |<ATC Part     |<Coating Type"

sSQLF = sSQLF & "|^Series                |^Case|^Value Range|^   |>ZHT  |^               |^Active|>CP     |^     "
If (Option4.value = True) Then
            sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER]," & _
                          "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                          "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11]," & _
                          "format([ACTIVE],'Yes/No')" & _
                  "FROM [TBL Power] WHERE [TRAY_ID]=" & TRAY_ID

    sSQL = sSQL & " AND [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ATC PART],[SERIES] ASC,[COATING]"

    sSQLF = "   |^TRAY_ID|^POWER_ID|<     |<ATC Part           "

    sSQLF = sSQLF & "   |<Coating Type|^Series                |^Case|^Value Range|^   |<POS 9                        |<POS 10                  |<POS 10                |POS 11 |^Active"

End If

If (Option5.value = True) Then
            sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER]," & _
                          "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                          "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11]," & _
                          "format([ACTIVE],'Yes/No')" & _
                  "FROM [TBL Power] WHERE [TRAY_ID]=" & TRAY_ID

    sSQL = sSQL & " AND [ACTIVE] = Yes  ORDER BY [ORDER]"

    sSQLF = "   |^TRAY_ID|^POWER_ID|<     |<ATC Part           "

    sSQLF = sSQLF & "   |<Coating Type|^Series                |^Case|^Value Range|^   |<POS 9                        |<POS 10                  |<POS 10                |POS 11 |^Active"

End If

Data6.RecordSource = sSQL
Data6.Refresh
 
MSFlexGrid6.FormatString = sSQLF

End Sub

Private Sub cmdUpdate_Click()

Data2.UpdateRecord
cmdRefresh1_Click

End Sub

Private Sub Command1_Click()

Data7.UpdateRecord
cmdRefresh6_Click

End Sub

Private Sub Command10_Click()

FindReverseLimit 2

End Sub

Private Sub Command11_Click()
txtLocation(1).Text = txtLocation(1).Text - 10
End Sub

Private Sub Command12_Click()
txtLocation(1).Text = txtLocation(1).Text + 10

End Sub

Private Sub Command13_Click()
txtLocation(1).Text = txtLocation(1).Text + 1
End Sub

Private Sub Command14_Click()
txtLocation(1).Text = txtLocation(1).Text - 1
End Sub

Private Sub Command15_Click()
txtLocation(1).Text = txtLocation(1).Text - 100
End Sub

Private Sub Command16_Click()
txtLocation(1).Text = txtLocation(1).Text + 100
End Sub

Private Sub Command2_Click()

Dim iSegment As Integer

iSegment = Val(TextSegment.Text)

  MoveToTarget 1, CLng(txtLocation(iSegment).Text)

End Sub

Private Sub Command3_Click()

MoveToTarget 2, CLng(lblCP.Caption)

End Sub

Private Sub Command4_Click()
FindReverseLimit 1
End Sub

Private Sub Command6_Click()

StopMotion 2

End Sub

Private Sub Command7_Click()
StopMotion 3
End Sub

Private Sub CommandAbort_Click()

StopMotion 1

'Initialize_Controller

End Sub

Private Sub CommandExcel_Click()

Dim objExcel As Object
Set objExcel = CreateObject("EXCEL.SHEET")
objExcel.Application.Visible = True

Screen.MousePointer = vbHourglass
 
Set FR_Database = OpenDatabase(ATC_LASER_BD)
Set TO_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TRAY_ID]                AS [SQL 1]," & _
              "[CASE]                   AS [SQL 3]," & _
              "[TITLE]                  AS [SQL 4]," & _
              "[ROWS]& ' X ' & [COLS]   AS [SQL 5]," & _
              "[ATC DWG]                AS [SQL 2]" & _
        "FROM  [TBL Tray Config] "

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount = 0) Then
    MsgBox "No Records"
    Exit Sub
End If

Dim iRow As Integer
iRow = iRow + 1
                ' (iRow,iCol)
objExcel.Application.Cells(1, 1).value = "TRAY_ID"
objExcel.Application.Cells(1, 2).value = "ATC DWG"
objExcel.Application.Cells(1, 3).value = "CASE"
objExcel.Application.Cells(1, 4).value = "TITLE"
objExcel.Application.Cells(1, 5).value = "Row X Col"

objExcel.Application.Cells(1, 7).value = "POWER_ID"
objExcel.Application.Cells(1, 8).value = "ATC PART"
objExcel.Application.Cells(1, 9).value = "COATING"
objExcel.Application.Cells(1, 10).value = "SERIES"
objExcel.Application.Cells(1, 11).value = "CASE"
objExcel.Application.Cells(1, 12).value = "VALUE"

iRow = iRow + 1
Do Until FR_Table.EOF
        objExcel.Application.Cells(iRow, 1).value = FR_Table.Fields("[SQL 1]")
        objExcel.Application.Cells(iRow, 2).value = FR_Table.Fields("[SQL 2]")
        objExcel.Application.Cells(iRow, 3).value = FR_Table.Fields("[SQL 3]")
        objExcel.Application.Cells(iRow, 4).value = FR_Table.Fields("[SQL 4]")
        objExcel.Application.Cells(iRow, 5).value = FR_Table.Fields("[SQL 5]")

        sSQL = "SELECT * FROM [TBL Power] " & _
               "WHERE [TRAY_ID] = " & FR_Table.Fields("[SQL 1]") & " AND [ACTIVE] = Yes  " & _
               "ORDER BY [TBL_ID]"
        
        Set TO_Table = TO_Database.OpenRecordset(sSQL)
        Do Until TO_Table.EOF
              objExcel.Application.Cells(iRow, 7).value = TO_Table.Fields("[TBL_ID]")
              objExcel.Application.Cells(iRow, 8).value = TO_Table.Fields("[ATC PART]")
              objExcel.Application.Cells(iRow, 9).value = TO_Table.Fields("[COATING]")
              objExcel.Application.Cells(iRow, 10).value = TO_Table.Fields("[SERIES]")
              objExcel.Application.Cells(iRow, 11).value = TO_Table.Fields("[CASE ]")
              objExcel.Application.Cells(iRow, 12).value = TO_Table.Fields("[VALUE]")
             TO_Table.MoveNext
             iRow = iRow + 1
        Loop
        FR_Table.MoveNext
Loop
 
TO_Database.Close
FR_Database.Close
                                                                                                     
Dim sFile As String
sFile = "C:\ATC\" & "TRAY LASER PARAMETERS.XLS"
                                                                
objExcel.SaveAs sFile
objExcel.Application.Quit
Set objExcel = Nothing
 
Screen.MousePointer = vbDefault
MsgBox "Tray Laser " & sFile, vbInformation, "Excel Format Download"

End Sub

Private Sub CommandInitialize_Click()

Screen.MousePointer = vbHourglass

Initialize_Controller

Load_Parameters

Screen.MousePointer = vbDefault

CommandInitialize.FontBold = True

End Sub

Private Sub CommandUpdateHeight_Click()

GlobalHeightUpdate

End Sub

Private Sub CommandUpdateRecord_Click()

Data3.UpdateRecord
cmdRefresh4_Click

End Sub

'/////////////////////////////////////////////////////////////////
' Form Load - Initializations
Private Sub Form_Load()

Caption = "Tray Configuration      " & ATC_DWG & "         " & ATC_VERSION

Dim sBuff As String

sBuff = "[1] X1 is the distance from Reverse Limit to Center of first Segment in Steps "
sBuff = sBuff & "[2] Calculate Segment 2,3,4 by Spacing * Inches to Steps * Motor Revolution"

cmdCalculate.ToolTipText = sBuff

Data1.DatabaseName = ATC_LASER_BD
Data2.DatabaseName = ATC_LASER_BD
Data3.DatabaseName = ATC_LASER_BD
Data6.DatabaseName = ATC_LASER_BD
Data7.DatabaseName = ATC_LASER_BD
Data8.DatabaseName = ATC_LASER_BD
                                
'MSFlexGrid4.Top = 0
MSFlexGrid4.Width = 5000
MSFlexGrid4.Height = 1200

MSFlexGrid1.Top = 0
MSFlexGrid1.Left = 0
MSFlexGrid1.Width = 13600
MSFlexGrid1.Height = 4000

MSFlexGrid6.Left = 0

cmdRefresh4_Click
MSFlexGrid4_Click
                
cmdRefresh1_Click
MSFlexGrid1_Click
 
Dim sSQL As String

Set FR_Database = OpenDatabase(ATC_LASER_BD)
sSQL = "SELECT * FROM  [TBL AXIS] WHERE [AXIS_ID] = 2"
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    vsbCP.Min = FR_Table.Fields("[RANGE MIN]")
    vsbCP.Max = FR_Table.Fields("[RANGE MAX]")
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Select Case OP_MODE
Case 0
        frmMain.Show
Case 1
        frmOPScreen.Show
End Select

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
TRAY_ID = Val(MSFlexGrid1.Text)
 
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10

Dim sSQL As String

Data2.DatabaseName = ATC_LASER_BD

sSQL = "SELECT * FROM  [TBL Tray Config] WHERE [TRAY_ID] = " & TRAY_ID
                    
Data2.RecordSource = sSQL
Data2.Refresh

cmdRefresh6_Click
MSFlexGrid6_Click

End Sub


Private Sub MSFlexGrid4_Click()

MSFlexGrid4.Col = 1
TBL_ID = Val(MSFlexGrid4.Text)
 
MSFlexGrid4.Col = 0
MSFlexGrid4.ColSel = MSFlexGrid4.Cols - 1 '10

Dim sSQL As String

sSQL = "SELECT * FROM [TBL Axis] WHERE [AXIS_ID] =" & TBL_ID

Data3.RecordSource = sSQL
Data3.Refresh

End Sub

Private Sub MSFlexGrid6_Click()

MSFlexGrid6.Col = 2
TBL_ID = Val(MSFlexGrid6.Text)
 
MSFlexGrid6.Col = 0
MSFlexGrid6.ColSel = MSFlexGrid6.Cols - 1 '10

Dim sSQL As String

sSQL = "SELECT * FROM  [TBL Power] WHERE [TBL_ID] = " & TBL_ID
                    
Data7.RecordSource = sSQL
Data7.Refresh

End Sub

Private Sub Option1_Click()
cmdRefresh6_Click
MSFlexGrid6_Click
End Sub

Private Sub Option2_Click()
cmdRefresh6_Click
MSFlexGrid6_Click
End Sub

Private Sub Option3_Click()
cmdRefresh6_Click
MSFlexGrid6_Click
End Sub

Private Sub Option4_Click()
cmdRefresh6_Click
MSFlexGrid6_Click
End Sub

Private Sub Option5_Click()
cmdRefresh6_Click
MSFlexGrid6_Click
End Sub

Private Sub Text11_LostFocus()
Text11.Text = UCase(Text11.Text)
End Sub

Private Sub Text22_GotFocus()
Text22.SelStart = 0
Text22.SelLength = Len(Text22)
End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5)
End Sub

Private Sub TextSegment_GotFocus()
TextSegment.SelStart = 0
TextSegment.SelLength = Len(TextSegment)
End Sub

Private Sub txtCameraPosition_GotFocus()

txtCameraPosition.SelStart = 0
txtCameraPosition.SelLength = Len(txtCameraPosition)

End Sub

Private Sub txtLocation_GotFocus(Index As Integer)
txtLocation(Index).SelStart = 0
txtLocation(Index).SelLength = Len(txtLocation(Index).Text)
End Sub

Private Sub vsbCP_Change()
lblCP.Caption = Format(vsbCP.value, "0")
End Sub
