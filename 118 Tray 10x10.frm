VERSION 5.00
Begin VB.Form frm10x10 
   Caption         =   "'B' Case 10 X 10 Carrier Tray"
   ClientHeight    =   11820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18765
   LinkTopic       =   "Form1"
   ScaleHeight     =   11820
   ScaleWidth      =   18765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommandAll 
      BackColor       =   &H00C0FFC0&
      Caption         =   "All"
      Height          =   300
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   7080
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      DataField       =   "CAMERA POS"
      DataSource      =   "Data2"
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
      Index           =   0
      Left            =   13200
      TabIndex        =   210
      Text            =   "CAMERA POS"
      ToolTipText     =   "CAMERA POS"
      Top             =   7440
      Width           =   1100
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rev Limit Tray"
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
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rev Limit Cam"
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
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   7320
      Width           =   1815
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
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   200
      Top             =   7920
      Width           =   1815
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
      TabIndex        =   193
      Top             =   8760
      Width           =   1935
      Begin VB.OptionButton optLogo5 
         Caption         =   "Abrasive"
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
         Left            =   240
         TabIndex        =   199
         Top             =   1860
         Width           =   1380
      End
      Begin VB.OptionButton optLogo 
         Caption         =   "Serialization"
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
         TabIndex        =   198
         Top             =   2235
         Width           =   1620
      End
      Begin VB.OptionButton optLogo4 
         Caption         =   "Omit"
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
         TabIndex        =   197
         Top             =   1485
         Width           =   900
      End
      Begin VB.OptionButton optLogo1 
         Caption         =   "Side"
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
         TabIndex        =   196
         Top             =   1110
         Width           =   900
      End
      Begin VB.OptionButton optLogo3 
         Caption         =   "Top"
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
         TabIndex        =   195
         Top             =   735
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optLogo2 
         Caption         =   "ATC"
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
         TabIndex        =   194
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.CommandButton CommandLoadDPSS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Load DPSS"
      Height          =   300
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtFrequency 
      BackColor       =   &H00C0FFC0&
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
      Left            =   11880
      TabIndex        =   182
      Text            =   "XXX"
      Top             =   9240
      Width           =   1095
   End
   Begin VB.TextBox txtMarkspeed 
      BackColor       =   &H00C0FFC0&
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
      Left            =   11880
      TabIndex        =   181
      Text            =   "XXXXXX"
      Top             =   9720
      Width           =   1095
   End
   Begin VB.TextBox txtPulseWidth 
      BackColor       =   &H00C0FFC0&
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
      Left            =   11880
      TabIndex        =   180
      Text            =   "X"
      ToolTipText     =   "PulseWidth"
      Top             =   10200
      Width           =   1095
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
      Height          =   360
      Left            =   13875
      TabIndex        =   179
      Top             =   9720
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
      Height          =   360
      Left            =   14670
      TabIndex        =   178
      Top             =   9720
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
      Height          =   360
      Left            =   15465
      TabIndex        =   177
      Top             =   9720
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
      Height          =   360
      Left            =   13080
      TabIndex        =   176
      Top             =   9720
      Width           =   800
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
      Height          =   360
      Left            =   13875
      TabIndex        =   175
      Top             =   9240
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
      Height          =   360
      Left            =   14670
      TabIndex        =   174
      Top             =   9240
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
      Height          =   360
      Left            =   15480
      TabIndex        =   173
      Top             =   9240
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
      Height          =   360
      Left            =   13080
      TabIndex        =   172
      Top             =   9240
      Width           =   800
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
      Height          =   360
      Left            =   13875
      TabIndex        =   171
      Top             =   10200
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
      Height          =   360
      Left            =   14670
      TabIndex        =   170
      Top             =   10200
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
      Height          =   360
      Left            =   15465
      TabIndex        =   169
      Top             =   10200
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
      Height          =   360
      Left            =   13080
      TabIndex        =   168
      Top             =   10200
      Width           =   800
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update Record"
      Height          =   300
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   167
      ToolTipText     =   "[FIXTURE] WHERE [MATRIX ID]"
      Top             =   10800
      Width           =   1485
   End
   Begin VB.Frame Frame6 
      Caption         =   " Laser Offset "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5280
      TabIndex        =   152
      Top             =   8760
      Width           =   4215
      Begin VB.TextBox txtOff 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         DataField       =   "OFFSET"
         DataSource      =   "Data2"
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
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   161
         Text            =   "0"
         ToolTipText     =   "OFFSET"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtXOffset 
         BackColor       =   &H00FFC0C0&
         DataField       =   "L X Offset"
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
         Left            =   600
         TabIndex        =   160
         Text            =   "0"
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtYOffset 
         BackColor       =   &H00FFC0C0&
         DataField       =   "L Y Offset"
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
         Left            =   2040
         TabIndex        =   159
         Text            =   "0"
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFC0C0&
         DataField       =   "L Y Offset 1"
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
         Index           =   10
         Left            =   2040
         TabIndex        =   158
         Text            =   "30.0"
         ToolTipText     =   "L Y Offset"
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFC0C0&
         DataField       =   "L X Offset 1"
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
         Index           =   9
         Left            =   600
         TabIndex        =   157
         Text            =   "30.0"
         ToolTipText     =   "L X Offset"
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton cmdU 
         Caption         =   "^"
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
         Left            =   3240
         TabIndex        =   156
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdD 
         Caption         =   "v"
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
         Left            =   3240
         TabIndex        =   155
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdR 
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
         Left            =   3480
         TabIndex        =   154
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdL 
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
         Left            =   3000
         TabIndex        =   153
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Off Set (0/1) "
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
         Left            =   480
         TabIndex        =   166
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label9 
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
         Left            =   1680
         TabIndex        =   165
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label8 
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
         TabIndex        =   164
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
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
         TabIndex        =   163
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
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
         Left            =   1680
         TabIndex        =   162
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CheckBox chkSegment 
      Caption         =   "Segments Inclusive"
      Height          =   255
      Left            =   8160
      TabIndex        =   151
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoadPosition 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load Position"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   360
      Width           =   1500
   End
   Begin VB.TextBox txtTargetPos 
      BackColor       =   &H00FFC0C0&
      DataField       =   "X 0"
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
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   149
      Text            =   "TP0"
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   " Tray Configuration "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14760
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton OptionTray8 
         Caption         =   "[8] 10 X 10 Chip"
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
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1980
      End
   End
   Begin VB.Data Data3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Data3 [TBL Size Location]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   6600
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.OptionButton OptionAll 
      Caption         =   " [All] Segments 1-10"
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
      Left            =   7800
      TabIndex        =   147
      Top             =   6240
      Width           =   2220
   End
   Begin VB.TextBox txtSegYDist 
      BackColor       =   &H00FFC0C0&
      DataField       =   "SEG Y DIST"
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
      Height          =   375
      Left            =   13320
      TabIndex        =   145
      Text            =   "30.0"
      Top             =   1080
      Width           =   735
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
      Height          =   375
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   7920
      Width           =   1935
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
      Height          =   375
      Left            =   13200
      MaxLength       =   10
      TabIndex        =   143
      Text            =   "Z"
      ToolTipText     =   "Z HEIGHT"
      Top             =   7920
      Width           =   1095
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
      Height          =   375
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   7920
      Width           =   495
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
      Height          =   375
      Left            =   15765
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   7920
      Width           =   495
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
      Height          =   375
      Left            =   15390
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   7920
      Width           =   375
   End
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
      Height          =   375
      Left            =   15015
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00FFC0C0&
      Caption         =   "To Segment"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   840
      Width           =   1500
   End
   Begin VB.TextBox txtTargetPos 
      BackColor       =   &H00FFC0C0&
      DataField       =   "X 4"
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
      Height          =   375
      Index           =   4
      Left            =   10320
      TabIndex        =   137
      Text            =   "TP4"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtTargetPos 
      BackColor       =   &H00FFC0C0&
      DataField       =   "X 3"
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
      Height          =   375
      Index           =   3
      Left            =   10320
      TabIndex        =   136
      Text            =   "TP3"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtTargetPos 
      BackColor       =   &H00FFC0C0&
      DataField       =   "X 2"
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
      Height          =   375
      Index           =   2
      Left            =   10320
      TabIndex        =   135
      Text            =   "TP2"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtTargetPos 
      BackColor       =   &H00FFC0C0&
      DataField       =   "X 1"
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
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   134
      Text            =   "TP1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtTargetPos 
      BackColor       =   &H00FFC0C0&
      DataField       =   "X 5"
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
      Height          =   375
      Index           =   5
      Left            =   10320
      TabIndex        =   133
      Text            =   "TP5"
      Top             =   5160
      Width           =   735
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   6240
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.CommandButton cmdPositionCamera 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Position Camera"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [TBL Tray Config]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   5880
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.OptionButton Option10 
      Caption         =   "[10] Segment 10 "
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
      Left            =   7800
      TabIndex        =   131
      Top             =   5535
      Width           =   1740
   End
   Begin VB.OptionButton Option9 
      Caption         =   "[9] Segment 9"
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
      Left            =   7800
      TabIndex        =   130
      Top             =   5160
      Width           =   1740
   End
   Begin VB.OptionButton Option8 
      Caption         =   "[8] Segment 8 "
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
      Left            =   7800
      TabIndex        =   129
      Top             =   4785
      Width           =   1740
   End
   Begin VB.OptionButton Option7 
      Caption         =   "[7] Segment 7"
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
      Left            =   7800
      TabIndex        =   128
      Top             =   4410
      Width           =   1740
   End
   Begin VB.OptionButton Option6 
      Caption         =   "[6] Segment 6 "
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
      Left            =   7800
      TabIndex        =   127
      Top             =   4035
      Width           =   1740
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5] Segment 5"
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
      Left            =   7800
      TabIndex        =   126
      Top             =   3660
      Width           =   1740
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4] Segment 4 "
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
      Left            =   7800
      TabIndex        =   125
      Top             =   3285
      Width           =   1740
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3] Segment 3 "
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
      Left            =   7800
      TabIndex        =   124
      Top             =   2910
      Width           =   1740
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2] Segment 2 "
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
      Left            =   7800
      TabIndex        =   123
      Top             =   2535
      Width           =   1740
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] Segment 1 "
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
      Left            =   7800
      TabIndex        =   122
      Top             =   2160
      Value           =   -1  'True
      Width           =   1740
   End
   Begin VB.Frame fraM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   6615
      Begin VB.Frame fraC 
         BackColor       =   &H000080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5880
         TabIndex        =   121
         Top             =   5520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   25
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   26
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   27
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   28
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   29
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   30
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   31
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   32
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   33
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   34
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   35
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   36
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   37
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   38
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   39
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   40
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   41
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   42
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   43
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   44
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   45
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   46
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   47
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   48
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   49
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   50
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   51
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   52
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   53
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   54
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   55
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   56
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   57
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   58
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   59
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   60
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   61
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   62
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   63
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   64
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   65
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   66
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   67
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   68
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   69
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   70
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   71
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   72
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   73
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   74
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   75
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   76
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   77
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   78
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   79
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   80
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   81
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   82
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   83
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   84
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   85
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   86
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   87
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   88
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   89
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   90
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   91
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   92
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   93
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   94
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   95
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   96
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   97
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   98
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton cmdMatrix 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   99
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5760
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdTrayConfig 
      Caption         =   "Tray Config"
      Height          =   300
      Left            =   13080
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdFire 
      BackColor       =   &H8000000A&
      Caption         =   "Fire"
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
      Left            =   10200
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1500
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
      Height          =   2175
      Left            =   2400
      TabIndex        =   7
      Top             =   8760
      Width           =   2655
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
         TabIndex        =   15
         Top             =   735
         Width           =   1245
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
         TabIndex        =   14
         Top             =   1110
         Width           =   1245
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
         TabIndex        =   13
         Top             =   1485
         Width           =   1245
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
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
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
         TabIndex        =   11
         Top             =   735
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
         TabIndex        =   10
         Top             =   1110
         Width           =   645
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
         TabIndex        =   9
         Top             =   1485
         Width           =   645
      End
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
         TabIndex        =   8
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
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
      Left            =   13080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1200
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse"
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
      Left            =   13080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H8000000A&
      Caption         =   "All"
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
      Left            =   11880
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1200
   End
   Begin VB.CommandButton cmdCorners 
      Caption         =   "Corners"
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
      Left            =   11880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdRow 
      Caption         =   "Row"
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
      Left            =   11880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1200
   End
   Begin VB.CommandButton cmdColumn 
      Caption         =   "Column"
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
      Left            =   13080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   " Array Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11760
      TabIndex        =   211
      Top             =   3960
      Width           =   4575
      Begin VB.TextBox TextXSPACE 
         BackColor       =   &H00FFC0C0&
         DataField       =   "XSPACE"
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
         Left            =   1320
         TabIndex        =   216
         Text            =   "0.18"
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton CommandInitArray 
         Caption         =   "Initialize and /Update"
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
         TabIndex        =   215
         Top             =   360
         Width           =   2580
      End
      Begin VB.TextBox TextYSPACE 
         BackColor       =   &H00FFC0C0&
         DataField       =   "YSPACE"
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
         Left            =   1320
         TabIndex        =   214
         Text            =   "0.18"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TextXOFF 
         BackColor       =   &H00FFC0C0&
         DataField       =   "XOFF"
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
         TabIndex        =   213
         Text            =   "2.25"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TextYOFF 
         BackColor       =   &H00FFC0C0&
         DataField       =   "YOFF"
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
         TabIndex        =   212
         Text            =   "2.25"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "X SPACE"
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
         Height          =   345
         Left            =   240
         TabIndex        =   220
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Y SPACE"
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
         Height          =   345
         Left            =   240
         TabIndex        =   219
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "X OFFSET"
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
         Height          =   345
         Left            =   2280
         TabIndex        =   218
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Y OFFSET"
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
         Height          =   345
         Left            =   2280
         TabIndex        =   217
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Label Label10 
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
      Left            =   15600
      TabIndex        =   209
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label LabelSIZE_LOC_ID 
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
      Left            =   17160
      TabIndex        =   208
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label11 
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
      Left            =   15600
      TabIndex        =   207
      Top             =   3000
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
      Left            =   17160
      TabIndex        =   206
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label13 
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
      Left            =   15600
      TabIndex        =   205
      Top             =   2520
      Width           =   1095
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
      Left            =   17160
      TabIndex        =   204
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label LabelFC 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FIRE_COUNT"
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
      Left            =   16680
      TabIndex        =   203
      Top             =   1680
      Width           =   885
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SERIES"
      DataField       =   "SERIES"
      DataSource      =   "Data2"
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
      Left            =   11640
      TabIndex        =   191
      Top             =   8640
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CEAB"
      DataField       =   "CASE SIZE"
      DataSource      =   "Data3"
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
      Left            =   13440
      TabIndex        =   190
      Top             =   8640
      Width           =   765
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   9960
      TabIndex        =   189
      Top             =   9240
      Width           =   1470
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   9960
      TabIndex        =   188
      Top             =   9720
      Width           =   1755
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   9960
      TabIndex        =   187
      Top             =   10200
      Width           =   1440
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CASE SIZE"
      DataField       =   "CASE NAME"
      DataSource      =   "Data3"
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
      Left            =   14400
      TabIndex        =   186
      Top             =   8640
      Width           =   1845
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "[0.02 to 250.0]"
      Height          =   195
      Left            =   10440
      TabIndex        =   185
      Top             =   9480
      Width           =   1035
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "[0 to 30000]"
      Height          =   195
      Left            =   10440
      TabIndex        =   184
      Top             =   9960
      Width           =   855
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "[2 to 65535]"
      Height          =   195
      Left            =   10440
      TabIndex        =   183
      Top             =   10440
      Width           =   855
   End
   Begin VB.Label lblSegment 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEGMENT #"
      Height          =   375
      Left            =   5160
      TabIndex        =   148
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Segment Y Distance"
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
      Height          =   495
      Left            =   11640
      TabIndex        =   146
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   2295
   End
End
Attribute VB_Name = "frm10x10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAll_Click(Index As Integer)

Dim X As Integer, Y As Integer, k As Integer
For X = 0 To 9
       For Y = 0 To 9
            k = X + (10 * Y)
                If (cmdMatrix(k).Enabled = True) Then
                    cmdMatrix(k).Caption = k
                    cmdMatrix(k).BackColor = &HC0FFFF
                End If
        Next Y
Next X


End Sub

Private Sub cmdColumn_Click(Index As Integer)

Dim Y As Integer
Dim i As Integer
 
For Y = 0 To 9
     Select Case Index
     Case 0
                If (cmdMatrix(Y).Caption <> "") Then
                    For i = 0 To 9
                        If cmdMatrix(i * 10 + Y).Enabled = True Then
                            cmdMatrix(i * 10 + Y).Caption = i * 10 + Y
                            cmdMatrix(i * 10 + Y).BackColor = &HC0FFFF
                        End If
                    Next i
                End If
     End Select
Next Y

End Sub

Private Sub cmdCorners_Click(Index As Integer)

cmdMatrix(0).Caption = 0
cmdMatrix(9).Caption = 9
cmdMatrix(90).Caption = 90
cmdMatrix(99).Caption = 99
cmdMatrix(0).BackColor = &HC0FFFF
cmdMatrix(9).BackColor = &HC0FFFF
cmdMatrix(90).BackColor = &HC0FFFF
cmdMatrix(99).BackColor = &HC0FFFF
                
End Sub


Private Sub cmdD_Click()
txtYOffset.Text = txtYOffset.Text - 0.001
End Sub


Private Sub cmdFire_Click()

 
    If (OptionAll.value = True) Then
         CommandAll_Click
         Exit Sub
    End If

    Select Case Load_Job
    Case 0
            Load_Job_From_File
    Case 1
            Select Case TRAY_MARK_ANGLE
            Case MARK_ANGLE_DEFAULT
                    Select Case MARK_ANGLE
                    Case MARK_ANGLE_DEFAULT
                             'ok
                    Case MARK_ANGLE_ROTATED
                             'set to default
                             Load_Job_From_File
                    End Select
            Case MARK_ANGLE_ROTATED
                    Select Case MARK_ANGLE
                    Case MARK_ANGLE_DEFAULT
                            'set to rotate
                            Load_Job_From_File
                    Case MARK_ANGLE_ROTATED
                            'ok
                    End Select
            End Select
    End Select
    Data1.UpdateRecord
    Data2.UpdateRecord
             
    CommandLoadDPSS_Click
             
    Select Case Val(txtOff.Text)
    Case 0
            TRAY_X_OFFSET = Val(txtXOffset.Text)
            TRAY_Y_OFFSET = Val(txtYOffset.Text)
    Case Else
            TRAY_X_OFFSET = Val(txtLocation(9).Text)
            TRAY_Y_OFFSET = Val(txtLocation(10).Text)
    End Select
    
    If OptLine(0).value = True Then
        MARK_MODE = 1
    End If
    If OptLine(1).value = True Then
        MARK_MODE = 2
    End If
    If OptLine(2).value = True Then
        MARK_MODE = 3
    End If
    If OptLine(3).value = True Then
        MARK_MODE = 4
    End If
    
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
       
    If (Option1.value = True) Then
        SEGMENT_ID = 1
    End If
    If (Option2.value = True) Then
        SEGMENT_ID = 2
    End If
    If (Option3.value = True) Then
        SEGMENT_ID = 3
    End If
    If (Option4.value = True) Then
        SEGMENT_ID = 4
    End If
    If (Option5.value = True) Then
        SEGMENT_ID = 5
    End If
    If (Option6.value = True) Then
        SEGMENT_ID = 6
    End If
    If (Option7.value = True) Then
        SEGMENT_ID = 7
    End If
    If (Option8.value = True) Then
        SEGMENT_ID = 8
    End If
    If (Option9.value = True) Then
        SEGMENT_ID = 9
    End If
    If (Option10.value = True) Then
        SEGMENT_ID = 10
    End If
    
    SEGMENTS_SELECT = 0
    If (chkSegment.value = vbChecked) Then
        SEGMENTS_SELECT = SEGMENT_ID
    End If
    
    Select Case SEGMENT_ID
    Case 1, 3, 5, 7, 9
            SEG_Y_DIST = 0
            SEG_Y_DIST_0 = 0
    Case Else
            SEG_Y_DIST = Val(txtSegYDist.Text)
            SEG_Y_DIST_1 = Val(txtSegYDist.Text)
    End Select
            
    Dim X As Integer, Y As Integer, k As Integer
    For X = 0 To 9
           For Y = 0 To 9
                k = X + (10 * Y)
                If (OptionAll.value = False) Then
                
                    If (cmdMatrix(k).Caption <> "") Then
                            FIRE_MATRIX(SEGMENT_ID, k) = 1
                    Else
                            FIRE_MATRIX(SEGMENT_ID, k) = 0
                    End If
                Else
                
                    If (cmdMatrix(k).Caption <> "") Then
                            FIRE_MATRIX(1, k) = 1
                    Else
                            FIRE_MATRIX(1, k) = 0
                    End If
                
                    FIRE_MATRIX(2, k) = 1
                    FIRE_MATRIX(3, k) = 1
                    FIRE_MATRIX(4, k) = 1
                    FIRE_MATRIX(5, k) = 1
                    FIRE_MATRIX(6, k) = 1
                    FIRE_MATRIX(7, k) = 1
                    FIRE_MATRIX(8, k) = 1
                    FIRE_MATRIX(9, k) = 1
                    FIRE_MATRIX(10, k) = 1
                End If
            Next Y
    Next X
                        
    cmdFire.Enabled = False
    cmdFire.BackColor = vbButtonFace
                        
    If (OptionAll.value = False) Then
            'FIRE SEGMENT
            Select Case SEGMENT_ID
            Case 1, 3, 5, 7, 9
                    SEG_Y_DIST = 0
                    SEG_Y_DIST_0 = 0
            Case Else
                    SEG_Y_DIST = Val(txtSegYDist.Text)
                    SEG_Y_DIST_1 = Val(txtSegYDist.Text)
            End Select
                        
            For X = 0 To 9
                   For Y = 0 To 9
                        k = X + (10 * Y)
                        If (FIRE_MATRIX(SEGMENT_ID, k) = 1) Then
                            Fire_Objects (k)
                            MARK_COUNT = MARK_COUNT + 1
                        End If
                    Next Y
            Next X

    End If
                
    '*********** PRODUCTION MODE
    Dim i As Integer

    If (OptionAll.value = True) Then
        '
        '  chg For SEGMENT_ID = 1 To SEGMENTS_SELECT
        '
        '
        For SEGMENT_ID = 1 To 10
            '        '
            'MOVE X AXIS TO NEXT SEGMENT
            '
            Select Case OP_MODE
            Case 0, 1
                    Select Case SEGMENT_ID
                    Case 1, 2
                            MoveToTarget 1, CLng(txtTargetPos(1).Text)
                    Case 3, 4
                            MoveToTarget 1, CLng(txtTargetPos(2).Text)
                    Case 5, 6
                            MoveToTarget 1, CLng(txtTargetPos(3).Text)
                    Case 7, 8
                            MoveToTarget 1, CLng(txtTargetPos(4).Text)
                    Case 9, 10
                            MoveToTarget 1, CLng(txtTargetPos(5).Text)
                    End Select
            End Select
            lblSegment.Caption = SEGMENT_ID
            lblSegment.Refresh
                          
            
            Select Case SEGMENT_ID
            Case 1, 3, 5, 7, 9
                    SEG_Y_DIST = 0
                    SEG_Y_DIST_0 = 0
            Case Else
                    SEG_Y_DIST = Val(txtSegYDist.Text)
                    SEG_Y_DIST_1 = Val(txtSegYDist.Text)
            End Select
            
            
            '
            'MOVE X AXIS TO NEXT SEGMENT
            '
            For X = 0 To 9
                   For Y = 0 To 9
                        k = X + (10 * Y)
                        If (FIRE_MATRIX(SEGMENT_ID, k) = 1) Then
                            Fire_Objects (k)
                            MARK_COUNT = MARK_COUNT + 1
                        End If
                    Next Y
            Next X
                     
            lblSegment.Caption = "Laser BUSY"
            lblSegment.Refresh
                        
        Next SEGMENT_ID
        
        MoveToTarget 1, CLng(txtTargetPos(0).Text)
                
    End If
         
    cmdFire.Enabled = True
    cmdFire.BackColor = vbGreen
    lblSegment.Caption = "Complete"
    lblSegment.Refresh
    LabelFC.Caption = FIRE_COUNT_ID
    WorkLogNew
End Sub

Private Sub cmdL_Click()
txtXOffset.Text = txtXOffset.Text - 0.001
End Sub

Private Sub cmdLoad_Click()

Screen.MousePointer = vbHourglass

Dim sFilename As String
sFilename = "C:\MARKER\JOB\ATC DPSS.WLJ"

Dim JobIndex As Long

AutomationInterface.LoadJobFromFile sFilename, JobIndex
AutomationInterface.GetObjCount ObjectCount
 
Load_Job = 1

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdLoadPosition_Click()
 
MoveToTarget 1, CLng(txtTargetPos(0).Text)

End Sub

Private Sub cmdMatrix_Click(Index As Integer)

If (cmdMatrix(Index).Caption = "") Then
    cmdMatrix(Index).Caption = Index
    cmdMatrix(Index).BackColor = &HC0FFFF
Else
    cmdMatrix(Index).Caption = ""
    cmdMatrix(Index).BackColor = &H808080
End If

End Sub

Private Sub cmdMatrix_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sBuff As String

Select Case Index
Case 0 To 100
        
        sBuff = Index & "  ( " & Format(gdLocation(XLOC, Index), "0.000") & " , " & Format(gdLocation(YLOC, Index), "0.000") & " ) "
                
        Label1.Caption = sBuff
                
End Select

End Sub

Private Sub cmdM3_Click()
    txtFrequency.Text = Val(txtFrequency.Text) - 0.01
    If Val(txtFrequency.Text) < 0.02 Then
        txtFrequency.Text = 0.02
    End If
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

Private Sub cmdMove_Click()

Dim SEGMENT_ID As Long

If (Option1.value = True) Then
        SEGMENT_ID = 1
End If
If (Option2.value = True) Then
        SEGMENT_ID = 1
End If
If (Option3.value = True) Then
        SEGMENT_ID = 2
End If
If (Option4.value = True) Then
        SEGMENT_ID = 2
End If
If (Option5.value = True) Then
        SEGMENT_ID = 3
End If
If (Option6.value = True) Then
        SEGMENT_ID = 3
End If
If (Option7.value = True) Then
        SEGMENT_ID = 4
End If
If (Option8.value = True) Then
        SEGMENT_ID = 4
End If
If (Option9.value = True) Then
        SEGMENT_ID = 5
End If
If (Option10.value = True) Then
        SEGMENT_ID = 5
End If

Select Case SEGMENT_ID
Case 1 To 5
        MoveToTarget 1, CLng(txtTargetPos(SEGMENT_ID).Text)
End Select

End Sub

Private Sub cmdMinus1_Click()
txtZHT.Text = txtZHT.Text + 1
End Sub

Private Sub cmdMinus11_Click()
txtZHT.Text = txtZHT.Text + 10
End Sub

Private Sub cmdPlus1_Click()
txtZHT.Text = txtZHT.Text - 1
End Sub

Private Sub cmdPlus11_Click()
txtZHT.Text = txtZHT.Text - 10
End Sub

Private Sub cmdPositionCamera_Click()

Dim sSQL As String

Set FR_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT [CAMERA POS] FROM [TBL Power] WHERE [TBL_ID] =" & POWER_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then

    MoveToTarget 2, CLng(FR_Table.Fields("[CAMERA POS]"))
    
End If
FR_Table.Close
FR_Database.Close

End Sub

Private Sub cmdR_Click()
txtXOffset.Text = txtXOffset.Text + 0.001
End Sub

Private Sub cmdReset_Click(Index As Integer)

Dim X As Integer, Y As Integer, k As Integer
For X = 0 To 9
       For Y = 0 To 9
            k = X + (10 * Y)
            
            Select Case Index
            Case 0
                    cmdMatrix(k).Caption = ""
                    cmdMatrix(k).BackColor = &H808080
            End Select
                        
        Next Y
Next X

End Sub

Private Sub cmdReverse_Click(Index As Integer)

Dim X As Integer, Y As Integer, k As Integer
For X = 0 To 9
       For Y = 0 To 9
            k = X + (10 * Y)
            
            Select Case Index
            Case 0
                    If (cmdMatrix(k).Enabled = True) Then
                        If (cmdMatrix(k).Caption = "") Then
                            cmdMatrix(k).Caption = k
                            cmdMatrix(k).BackColor = &HC0FFFF
                        Else
                            cmdMatrix(k).Caption = ""
                            cmdMatrix(k).BackColor = &H808080
                        End If
                    End If
            End Select

        Next Y
Next X

End Sub

Private Sub cmdRow_Click(Index As Integer)

Dim Y As Integer, k As Integer
Dim i As Integer
 
For Y = 0 To 9
     k = (10 * Y)
     Select Case Index
     Case 0
                If (cmdMatrix(k).Caption <> "") Then
                    For i = 0 To 9
                         cmdMatrix(i + k).Caption = i + k
                         cmdMatrix(i + k).BackColor = &HC0FFFF
                    Next i
                End If
   
     End Select

Next Y
 
End Sub

Private Sub cmdTrayConfig_Click()

Dim X As Integer, Y As Integer, k As Integer

For X = 0 To 9
       For Y = 0 To 9
            k = X + (10 * Y)
            cmdMatrix(k).Enabled = True
            cmdMatrix(k).BackColor = &H808080
        Next Y
Next X

End Sub

Private Sub cmdU_Click()
txtYOffset.Text = txtYOffset.Text + 0.001
End Sub

Private Sub cmdZHT_Click()
MoveToTarget 3, CLng(txtZHT.Text)
End Sub

Private Sub Command1_Click()
FindReverseLimit 2
 
End Sub

Private Sub Command2_Click()
FindReverseLimit 3
 
End Sub

Private Sub Command3_Click()
FindReverseLimit 1
End Sub

Private Sub CommandAll_Click()

    Select Case Load_Job
    Case 0
            Load_Job_From_File
    Case 1
            Select Case TRAY_MARK_ANGLE
            Case MARK_ANGLE_DEFAULT
                    Select Case MARK_ANGLE
                    Case MARK_ANGLE_DEFAULT
                             'ok
                    Case MARK_ANGLE_ROTATED
                             'set to default
                             Load_Job_From_File
                    End Select
            Case MARK_ANGLE_ROTATED
                    Select Case MARK_ANGLE
                    Case MARK_ANGLE_DEFAULT
                            'set to rotate
                            Load_Job_From_File
                    Case MARK_ANGLE_ROTATED
                            'ok
                    End Select
            End Select
    End Select
    Data1.UpdateRecord
    Data2.UpdateRecord
             
    CommandLoadDPSS_Click
             
    Select Case Val(txtOff.Text)
    Case 0
            TRAY_X_OFFSET = Val(txtXOffset.Text)
            TRAY_Y_OFFSET = Val(txtYOffset.Text)
    Case Else
            TRAY_X_OFFSET = Val(txtLocation(9).Text)
            TRAY_Y_OFFSET = Val(txtLocation(10).Text)
    End Select
        
    If OptLine(0).value = True Then
        MARK_MODE = 1
    End If
    If OptLine(1).value = True Then
        MARK_MODE = 2
    End If
    If OptLine(2).value = True Then
        MARK_MODE = 3
    End If
    If OptLine(3).value = True Then
        MARK_MODE = 4
    End If
    
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
       
    ' SAME OPENING
        
    For SEGMENT_ID = 1 To 10
    
            SEGMENTS_SELECT = 0
            Select Case SEGMENT_ID
            Case 1, 3, 5, 7, 9
                    SEG_Y_DIST = 0
                    SEG_Y_DIST_0 = 0
            Case Else
                    SEG_Y_DIST = Val(txtSegYDist.Text)
                    SEG_Y_DIST_1 = Val(txtSegYDist.Text)
            End Select
                    
            Dim X As Integer, Y As Integer, k As Integer
            For X = 0 To 9
                   For Y = 0 To 9
                        k = X + (10 * Y)
                        If (OptionAll.value = False) Then
                        
                            If (cmdMatrix(k).Caption <> "") Then
                                    FIRE_MATRIX(SEGMENT_ID, k) = 1
                            Else
                                    FIRE_MATRIX(SEGMENT_ID, k) = 0
                            End If
                        Else
                        
                            If (cmdMatrix(k).Caption <> "") Then
                                    FIRE_MATRIX(1, k) = 1
                            Else
                                    FIRE_MATRIX(1, k) = 0
                            End If
                        
                            FIRE_MATRIX(2, k) = 1
                            FIRE_MATRIX(3, k) = 1
                            FIRE_MATRIX(4, k) = 1
                            FIRE_MATRIX(5, k) = 1
                            FIRE_MATRIX(6, k) = 1
                            FIRE_MATRIX(7, k) = 1
                            FIRE_MATRIX(8, k) = 1
                            FIRE_MATRIX(9, k) = 1
                            FIRE_MATRIX(10, k) = 1
                        End If
                    Next Y
            Next X
                        
    Next SEGMENT_ID
                        
    cmdFire.Enabled = False
    cmdFire.BackColor = vbButtonFace
                        

                
         
    cmdFire.Enabled = True
    cmdFire.BackColor = vbGreen
    lblSegment.Caption = "Complete"
    lblSegment.Refresh
    LabelFC.Caption = FIRE_COUNT_ID
    WorkLogNew


End Sub

Private Sub CommandInitArray_Click()

Dim X As Integer, Y As Integer, k As Integer
'
'    OFFSET TO POSITION (1,1)
'
Dim dOffSet(1) As Double

'
' TRAY B CASE 10 X 10
'
'104 Tray Laser chg 12/06/2012 Array 10x10 0.181

For X = 0 To 9
       For Y = 0 To 9
            k = X + (10 * Y)
            gdLocation(XLOC, k) = X * Val(TextXSPACE.Text)
            gdLocation(YLOC, k) = Y * Val(TextYSPACE.Text)
        Next Y
Next X

dOffSet(XLOC) = Val(TextXOFF.Text)
dOffSet(YLOC) = Val(TextYOFF.Text)
           
' 9 * 0.5 / 2
' ADD OFFSET ADJUST
'
For k = 0 To 99
    gdLocation(XLOC, k) = gdLocation(XLOC, k) - dOffSet(XLOC)
    gdLocation(YLOC, k) = -(gdLocation(YLOC, k) - dOffSet(YLOC))
Next k

End Sub

Private Sub CommandLoadDPSS_Click()
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

Select Case MARK_MODE
Case 1
        LASER_TXT1 = Text1.Text
Case 2
        LASER_TXT1 = Text1.Text
        LASER_TXT2 = Text2.Text
Case 3
        LASER_TXT1 = Text1.Text
        LASER_TXT2 = Text2.Text
        LASER_TXT3 = Text3.Text
Case 4
        LASER_TXT1 = Text1.Text
        LASER_TXT2 = Text2.Text
        LASER_TXT3 = Text3.Text
        LASER_TXT4 = Text4.Text
End Select

LoadDPSS
End Sub

Private Sub Form_Activate()

Select Case INITIALIZE_TRAY
Case 1
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
         
        cmdPositionCamera_Click
        cmdZHT_Click
        cmdLoadPosition_Click
        Initialize_Fire_Matrix
        Me.Enabled = True
         
        optLogo1.Enabled = False
        optLogo4.Enabled = False
        Screen.MousePointer = vbDefault
End Select
LabelFC.Caption = FIRE_COUNT_ID

CommandInitArray_Click

End Sub

Private Sub Form_Load()

Caption = "ATC " & Get_Title & "   " & ATC_DWG & "         " & ATC_VERSION

LabelSIZE_LOC_ID.Caption = SIZE_LOC_ID
LabelTRAY_ID.Caption = TRAY_ID
LabelPOWER_ID.Caption = POWER_ID

cmdFire.BackColor = vbGreen
Init_Array (MATRIX_ID)

Dim X As Integer, Y As Integer, k As Integer
For X = 0 To 9
       For Y = 0 To 9
            k = X + (10 * Y)
            cmdMatrix(k).Caption = ""
        Next Y
Next X

If Len(Text1.Text & "X") = 1 Then
    Text1.Text = LASER_TXT1
End If
If Len(Text2.Text & "X") = 1 Then
    Text2.Text = LASER_TXT2
End If
If Len(Text3.Text & "X") = 1 Then
    Text3.Text = LASER_TXT3
End If
If Len(Text4.Text & "X") = 1 Then
    Text4.Text = LASER_TXT4
End If

cmdTrayConfig_Click

Data1.DatabaseName = ATC_LASER_BD
Data2.DatabaseName = ATC_LASER_BD
Data3.DatabaseName = ATC_LASER_BD
 
Dim sSQL As String

sSQL = "SELECT * FROM  [TBL Tray Config] WHERE [TRAY_ID]= " & TRAY_ID
Data1.RecordSource = sSQL
Data1.Refresh

sSQL = "SELECT * FROM [TBL Power] WHERE [TBL_ID] =" & POWER_ID
Data2.RecordSource = sSQL
Data2.Refresh

sSQL = "SELECT * FROM  [TBL SIZE LOC] WHERE [SIZE_LOC_ID] = " & SIZE_LOC_ID

Data3.RecordSource = sSQL
Data3.Refresh


End Sub

Sub Init_Array(iMS As Integer)

Dim X As Integer, Y As Integer, k As Integer
'
'    OFFSET TO POSITION (1,1)
'
Dim dOffSet(1) As Double

'
' TRAY B CASE 10 X 10
'
'104 Tray Laser chg 12/06/2012 Array 10x10 0.181

For X = 0 To 9
       For Y = 0 To 9
            k = X + (10 * Y)
            gdLocation(XLOC, k) = X * 0.18
            gdLocation(YLOC, k) = Y * 0.18
        Next Y
Next X

dOffSet(XLOC) = 9 * 0.5 / 2
dOffSet(YLOC) = 9 * 0.5 / 2
           
'
' ADD OFFSET ADJUST
'
For k = 0 To 99
    gdLocation(XLOC, k) = gdLocation(XLOC, k) - dOffSet(XLOC)
    gdLocation(YLOC, k) = -(gdLocation(YLOC, k) - dOffSet(YLOC))
Next k
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

Select Case OP_MODE
Case 0
        frmMain.Show
Case 1
        frmOPScreen.Show
End Select

End Sub

Private Sub Option1_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option10_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option2_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option3_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option4_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option5_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option6_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option7_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option8_Click()
cmdTrayConfig_Click
End Sub

Private Sub Option9_Click()
cmdTrayConfig_Click
End Sub

Private Sub OptionAll_Click()
cmdTrayConfig_Click
End Sub

Private Sub txtScale_GotFocus(Index As Integer)
txtScale(Index).SelStart = 0
txtScale(Index).SelLength = Len(txtScale(Index))
End Sub

