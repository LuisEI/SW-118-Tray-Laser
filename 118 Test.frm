VERSION 5.00
Begin VB.Form frmTestScreen 
   Caption         =   "111 DPPS Laser Test Panel "
   ClientHeight    =   13245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18660
   Icon            =   "118 Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   13245
   ScaleWidth      =   18660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextLOCATION 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   10680
      TabIndex        =   186
      Text            =   "LS"
      Top             =   10080
      Width           =   495
   End
   Begin VB.CommandButton CommandSet 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Set"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   10680
      Width           =   795
   End
   Begin VB.TextBox TextCOUNT 
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
      Left            =   6360
      TabIndex        =   184
      Text            =   "0"
      ToolTipText     =   "Total time"
      Top             =   10560
      Width           =   600
   End
   Begin VB.CommandButton CommandExit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit to Main"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   12000
      Width           =   2820
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Options "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   600
      TabIndex        =   173
      Top             =   5640
      Width           =   3015
      Begin VB.CheckBox Check3 
         Caption         =   "[3] Error 2"
         Height          =   255
         Left            =   360
         TabIndex        =   182
         Top             =   1110
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "[2] Error 1"
         Height          =   255
         Left            =   360
         TabIndex        =   181
         Top             =   735
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "[1] Part Present"
         Height          =   255
         Left            =   360
         TabIndex        =   180
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Caption         =   "[4] Mark Position Right"
         Height          =   255
         Left            =   360
         TabIndex        =   179
         Top             =   1485
         Width           =   2055
      End
      Begin VB.CheckBox Check5 
         Caption         =   "[5] Mark Position Left"
         Height          =   255
         Left            =   360
         TabIndex        =   178
         Top             =   1860
         Width           =   2175
      End
      Begin VB.CheckBox Check6 
         Caption         =   "[6] Stop 1"
         Height          =   255
         Left            =   360
         TabIndex        =   177
         Top             =   2235
         Width           =   2175
      End
      Begin VB.CheckBox Check7 
         Caption         =   "[7] Stop 2"
         Height          =   255
         Left            =   360
         TabIndex        =   176
         Top             =   2610
         Width           =   2055
      End
      Begin VB.CheckBox Check8 
         Caption         =   "[8] Step Count +1"
         Height          =   255
         Left            =   360
         TabIndex        =   175
         Top             =   2985
         Width           =   2055
      End
      Begin VB.CheckBox Check9 
         Caption         =   "[9] Interlocks"
         Height          =   255
         Left            =   360
         TabIndex        =   174
         Top             =   3360
         Width           =   2055
      End
   End
   Begin VB.Frame fraPTMode 
      Caption         =   " Run Mode "
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
      Height          =   855
      Left            =   6000
      TabIndex        =   170
      Top             =   240
      Width           =   4095
      Begin VB.CommandButton CommandP 
         Caption         =   "Prod"
         Height          =   300
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   360
         Width           =   1600
      End
      Begin VB.CommandButton CommandT 
         Caption         =   "Test"
         Height          =   300
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   360
         Width           =   1600
      End
   End
   Begin VB.CommandButton CommandConfiguration 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Configuration"
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
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   1680
      Width           =   2820
   End
   Begin VB.CommandButton CommandResetFireCount 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reset Fire Count"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   11400
      Width           =   1755
   End
   Begin VB.CommandButton CommandProductionRun1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Production Run 1"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10440
      Width           =   2820
   End
   Begin VB.CommandButton CommandProductionRun 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Production Run"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9960
      Width           =   2820
   End
   Begin VB.OptionButton OptionBoth 
      Caption         =   "Both"
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
      TabIndex        =   10
      Top             =   5160
      Width           =   885
   End
   Begin VB.CommandButton cmdResetPosition 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reset Position"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   10920
      Width           =   1755
   End
   Begin VB.CommandButton cmdResetCount 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reset Count"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   10560
      Width           =   1755
   End
   Begin VB.Data Data2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Data2 FROM [FIXTURE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\111 MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL CONFIG"
      Top             =   12360
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "POSITION 2"
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
      Left            =   9240
      TabIndex        =   67
      ToolTipText     =   "Total time"
      Top             =   10920
      Width           =   720
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "POSITION 1"
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
      Left            =   9240
      TabIndex        =   66
      ToolTipText     =   "Total time"
      Top             =   10560
      Width           =   720
   End
   Begin VB.TextBox txtPause 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PAUSE"
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
      Left            =   6360
      TabIndex        =   153
      Text            =   "0.0"
      Top             =   12000
      Width           =   615
   End
   Begin VB.OptionButton optMarkRight 
      Caption         =   "Right"
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
      TabIndex        =   9
      Top             =   5160
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.OptionButton optMarkLeft 
      Caption         =   "Left"
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
      Left            =   720
      TabIndex        =   8
      Top             =   5160
      Width           =   885
   End
   Begin VB.Frame FrameHandler 
      Caption         =   " Handler "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   152
      Top             =   720
      Width           =   4335
      Begin VB.OptionButton optHand 
         Caption         =   "1 Left Side (A)"
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1600
      End
      Begin VB.OptionButton optHand 
         Caption         =   "2 Right Side (B)"
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
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      DataField       =   "CASE SIZE"
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
      Left            =   10800
      TabIndex        =   148
      Text            =   "Z"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      DataField       =   "CAPTION"
      DataSource      =   "Data1"
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
      Left            =   7920
      TabIndex        =   147
      Text            =   "Text1"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdProfile 
      BackColor       =   &H00C0FFFF&
      Caption         =   "COM Profile"
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
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2820
   End
   Begin VB.CommandButton cmdPower 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Power && Scale Factor"
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
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   2820
   End
   Begin VB.CommandButton cmdFixture 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fixture Select"
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
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   2820
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Load Job"
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
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2820
   End
   Begin VB.CommandButton cmdSetObj 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SetObject"
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
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2820
   End
   Begin VB.CommandButton cmdFireDPPS 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SetObjPos, MarkObj"
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
      Left            =   720
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   2820
   End
   Begin VB.Data Data1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Data1 FROM [FIXTURE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\111 MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FIXTURE"
      Top             =   12000
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Frame FramePower 
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
      Height          =   2415
      Left            =   5880
      TabIndex        =   140
      Top             =   4320
      Width           =   6735
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Update Record"
         Height          =   300
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   1920
         Width           =   1725
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
         Left            =   2400
         TabIndex        =   39
         Text            =   "XXX"
         ToolTipText     =   "[0.02 to 250.0]"
         Top             =   1440
         Width           =   855
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
         Left            =   2400
         TabIndex        =   29
         Text            =   "XXXXXX"
         ToolTipText     =   "[0 to 30000]"
         Top             =   480
         Width           =   855
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
         Left            =   2400
         TabIndex        =   34
         Text            =   "X"
         ToolTipText     =   "PulseWidth [2 to 65535]"
         Top             =   960
         Width           =   855
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
         TabIndex        =   30
         Top             =   480
         Width           =   700
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
         Left            =   5715
         TabIndex        =   33
         Top             =   480
         Width           =   700
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
         Left            =   5010
         TabIndex        =   32
         Top             =   480
         Width           =   700
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
         Left            =   4305
         TabIndex        =   31
         Top             =   480
         Width           =   700
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
         Left            =   4305
         TabIndex        =   36
         Top             =   960
         Width           =   700
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
         Left            =   5010
         TabIndex        =   37
         Top             =   960
         Width           =   700
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
         Left            =   5715
         TabIndex        =   38
         Top             =   960
         Width           =   700
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
         TabIndex        =   35
         Top             =   960
         Width           =   700
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
         Left            =   4305
         TabIndex        =   41
         Top             =   1440
         Width           =   700
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
         Left            =   5010
         TabIndex        =   42
         Top             =   1440
         Width           =   700
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
         Left            =   5715
         TabIndex        =   141
         Top             =   1440
         Width           =   700
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
         TabIndex        =   40
         Top             =   1440
         Width           =   700
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   145
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   360
         TabIndex        =   144
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label8 
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
         Left            =   360
         TabIndex        =   143
         Top             =   960
         Width           =   1695
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
      Height          =   2295
      Left            =   6000
      TabIndex        =   138
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton optLogo4 
         Caption         =   "[4] ATC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   300
         TabIndex        =   19
         Top             =   1740
         Width           =   1140
      End
      Begin VB.OptionButton optLogo2 
         Caption         =   "[2] Top"
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
         Height          =   300
         Left            =   300
         TabIndex        =   17
         Top             =   900
         Width           =   1380
      End
      Begin VB.OptionButton optLogo1 
         Caption         =   "[1] Side"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   300
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton optLogo3 
         Caption         =   "[3] Omit"
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
         Height          =   300
         Left            =   300
         TabIndex        =   18
         Top             =   1320
         Width           =   1380
      End
      Begin VB.OptionButton optLogo6 
         Caption         =   "[6] Abrasive"
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
         Left            =   300
         TabIndex        =   139
         Top             =   2655
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.OptionButton optLogo5 
         Caption         =   "[5] Serialization"
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
         Left            =   300
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
         Width           =   2100
      End
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
      Left            =   9360
      TabIndex        =   137
      Top             =   1200
      Width           =   3255
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
         Left            =   1320
         TabIndex        =   22
         Text            =   "AAAA"
         Top             =   480
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
         TabIndex        =   27
         Top             =   1605
         Visible         =   0   'False
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
         TabIndex        =   25
         Top             =   1230
         Visible         =   0   'False
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
         TabIndex        =   23
         Top             =   855
         Visible         =   0   'False
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
         TabIndex        =   21
         Top             =   480
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
         Left            =   1320
         TabIndex        =   28
         Text            =   "DDDD"
         Top             =   1605
         Visible         =   0   'False
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
         Left            =   1320
         TabIndex        =   26
         Text            =   "CCCC"
         Top             =   1230
         Visible         =   0   'False
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
         Left            =   1320
         TabIndex        =   24
         Text            =   "BBBB"
         Top             =   855
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Stop"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   11400
      Width           =   2820
   End
   Begin VB.CommandButton cmdProductionClear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Clear"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10920
      Width           =   2820
   End
   Begin VB.CommandButton cmdProductionRun 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Test Run"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9480
      Width           =   2820
   End
   Begin VB.Frame Frame12 
      Caption         =   " Inputs "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   13080
      TabIndex        =   109
      Top             =   120
      Width           =   5415
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   2
         Left            =   3960
         TabIndex        =   133
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   7
         Left            =   3960
         TabIndex        =   132
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   6
         Left            =   3960
         TabIndex        =   131
         Top             =   2580
         Width           =   1200
      End
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   5
         Left            =   3960
         TabIndex        =   130
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   4
         Left            =   3960
         TabIndex        =   129
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   3
         Left            =   3960
         TabIndex        =   128
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   1
         Left            =   3960
         TabIndex        =   127
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   7
         Left            =   3240
         TabIndex        =   126
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   6
         Left            =   3240
         TabIndex        =   125
         Top             =   2580
         Width           =   495
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   5
         Left            =   3240
         TabIndex        =   124
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   4
         Left            =   3240
         TabIndex        =   123
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   3
         Left            =   3240
         TabIndex        =   122
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   2
         Left            =   3240
         TabIndex        =   121
         Top             =   900
         Width           =   495
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   1
         Left            =   3240
         TabIndex        =   120
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblInput 
         Caption         =   "3"
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
         Index           =   3
         Left            =   240
         TabIndex        =   119
         Top             =   1320
         Width           =   2600
      End
      Begin VB.Label lblInput 
         Caption         =   "7"
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
         Index           =   7
         Left            =   240
         TabIndex        =   118
         Top             =   3000
         Width           =   2600
      End
      Begin VB.Label lblInput 
         Caption         =   "6"
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
         Left            =   240
         TabIndex        =   117
         Top             =   2580
         Width           =   2600
      End
      Begin VB.Label lblInput 
         Caption         =   "2"
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
         Index           =   2
         Left            =   240
         TabIndex        =   116
         Top             =   900
         Width           =   2600
      End
      Begin VB.Label lblInput 
         Caption         =   "5"
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
         Left            =   240
         TabIndex        =   115
         Top             =   2160
         Width           =   2600
      End
      Begin VB.Label lblInput 
         Caption         =   "4"
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
         Index           =   4
         Left            =   240
         TabIndex        =   114
         Top             =   1740
         Width           =   2600
      End
      Begin VB.Label lblInput 
         Caption         =   "1"
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
         Left            =   240
         TabIndex        =   113
         Top             =   480
         Width           =   2600
      End
      Begin VB.Label lblInput 
         Caption         =   "8"
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
         Index           =   8
         Left            =   240
         TabIndex        =   112
         Top             =   3480
         Width           =   2595
      End
      Begin VB.Label lblInp 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   8
         Left            =   3240
         TabIndex        =   111
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblInpV 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   420
         Index           =   8
         Left            =   3960
         TabIndex        =   110
         Top             =   3480
         Width           =   1200
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   " Ouptuts [1:0] [Off/On]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   13080
      TabIndex        =   100
      Top             =   4320
      Width           =   5415
      Begin VB.CommandButton cmdOut 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   720
         TabIndex        =   74
         Top             =   2880
         Width           =   2000
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   720
         TabIndex        =   73
         Top             =   2460
         Width           =   2000
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   720
         TabIndex        =   72
         Top             =   2040
         Width           =   2000
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   720
         TabIndex        =   71
         Top             =   1620
         Width           =   2000
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   720
         TabIndex        =   70
         Top             =   1200
         Width           =   2000
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   720
         TabIndex        =   69
         Top             =   780
         Width           =   2000
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   720
         TabIndex        =   68
         Top             =   360
         Width           =   2000
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   8
         Left            =   720
         TabIndex        =   75
         Top             =   3360
         Width           =   2000
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
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
         Index           =   7
         Left            =   3240
         TabIndex        =   108
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
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
         Left            =   3240
         TabIndex        =   107
         Top             =   2460
         Width           =   600
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
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
         Left            =   3240
         TabIndex        =   106
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
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
         Index           =   4
         Left            =   3240
         TabIndex        =   105
         Top             =   1620
         Width           =   600
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
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
         Index           =   3
         Left            =   3240
         TabIndex        =   104
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
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
         Index           =   2
         Left            =   3240
         TabIndex        =   103
         Top             =   780
         Width           =   600
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
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
         Left            =   3240
         TabIndex        =   102
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
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
         Index           =   8
         Left            =   3240
         TabIndex        =   101
         Top             =   3360
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Motion Index "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   13080
      TabIndex        =   94
      Top             =   8400
      Width           =   5415
      Begin VB.CommandButton CommandPULSE5 
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
         Left            =   2160
         TabIndex        =   78
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton CommandPULSE8 
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
         Left            =   3405
         TabIndex        =   81
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton CommandPULSE7 
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
         Left            =   3030
         TabIndex        =   80
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton CommandPULSE6 
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
         Left            =   2655
         TabIndex        =   79
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton CommandPULSE2 
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
         Left            =   2655
         TabIndex        =   85
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton CommandPULSE3 
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
         Left            =   3030
         TabIndex        =   86
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton CommandPULSE4 
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
         Left            =   3405
         TabIndex        =   87
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CommandPULSE1 
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
         Left            =   2160
         TabIndex        =   84
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdLoadParameters 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Load Parameters"
         Height          =   540
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2880
         Width           =   1155
      End
      Begin VB.TextBox txtDecel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "DECELERATION"
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
         Left            =   3000
         TabIndex        =   91
         Text            =   "30.0"
         Top             =   3000
         Width           =   675
      End
      Begin VB.TextBox txtAccel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ACCELERATION"
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
         Left            =   3000
         TabIndex        =   90
         Text            =   "30.0"
         Top             =   2640
         Width           =   675
      End
      Begin VB.TextBox txtVel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "VELOCITY"
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
         Left            =   3000
         TabIndex        =   89
         Text            =   "100.0"
         Top             =   2280
         Width           =   675
      End
      Begin VB.CommandButton cmdMovePocket 
         Caption         =   "Index Pocket"
         Height          =   300
         Left            =   3000
         TabIndex        =   99
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdIndex2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pulse [1 - 50]"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1680
         Width           =   1755
      End
      Begin VB.CommandButton cmdIndex 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pulse [1-320]"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdMove 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Index Pocket"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label Label17 
         Caption         =   "Deceleration (RPS/S)"
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
         Left            =   360
         TabIndex        =   162
         Top             =   3000
         Width           =   2400
      End
      Begin VB.Label Label16 
         Caption         =   "Velocity (RPM)"
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
         Left            =   360
         TabIndex        =   161
         Top             =   2280
         Width           =   2400
      End
      Begin VB.Label Label4 
         Caption         =   "Acceleration (RPS/S)"
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
         Left            =   360
         TabIndex        =   160
         Top             =   2640
         Width           =   2400
      End
      Begin VB.Label lblIndex2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   4320
         TabIndex        =   88
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblIndex 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   4320
         TabIndex        =   82
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame fraRight 
      Caption         =   "Mark Position Right"
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
      Left            =   8520
      TabIndex        =   97
      Top             =   6840
      Width           =   4095
      Begin VB.CommandButton CommandUP11 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   60
         Top             =   855
         Width           =   500
      End
      Begin VB.CommandButton CommandUP22 
         Caption         =   "^^"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   59
         Top             =   480
         Width           =   500
      End
      Begin VB.CommandButton CommandDN11 
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
         TabIndex        =   61
         Top             =   1230
         Width           =   500
      End
      Begin VB.CommandButton CommandDN22 
         Caption         =   "vv"
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
         TabIndex        =   62
         Top             =   1605
         Width           =   500
      End
      Begin VB.CommandButton cmdPlus11 
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
         Left            =   240
         TabIndex        =   54
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMinus11 
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
         Left            =   1485
         TabIndex        =   57
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMinus1 
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
         Left            =   1110
         TabIndex        =   56
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdPlus1 
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
         Left            =   735
         TabIndex        =   55
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtOffsetY 
         Alignment       =   1  'Right Justify
         DataField       =   "Y OFFSET"
         DataSource      =   "Data1"
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
         Left            =   2280
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtOffsetX 
         Alignment       =   1  'Right Justify
         DataField       =   "X OFFSET"
         DataSource      =   "Data1"
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
         Left            =   720
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Y"
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
         Left            =   1800
         TabIndex        =   159
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "X"
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
         Left            =   240
         TabIndex        =   158
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.Frame fraLeft 
      Caption         =   "Mark Position Left"
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
      Left            =   4440
      TabIndex        =   95
      Top             =   6840
      Width           =   3975
      Begin VB.CommandButton CommandUP1 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   50
         Top             =   840
         Width           =   500
      End
      Begin VB.CommandButton CommandUP2 
         Caption         =   "^^"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   49
         Top             =   480
         Width           =   500
      End
      Begin VB.CommandButton CommandDN1 
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
         TabIndex        =   51
         Top             =   1230
         Width           =   500
      End
      Begin VB.CommandButton CommandDN2 
         Caption         =   "vv"
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
         TabIndex        =   52
         Top             =   1605
         Width           =   500
      End
      Begin VB.CommandButton cmdPlus22 
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
         Left            =   240
         TabIndex        =   44
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMinus22 
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
         Left            =   1485
         TabIndex        =   47
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMinus2 
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
         Left            =   1110
         TabIndex        =   46
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdPlus2 
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
         Left            =   735
         TabIndex        =   45
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtOffsetX2 
         Alignment       =   1  'Right Justify
         DataField       =   "X OFFSET 2"
         DataSource      =   "Data1"
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
         Left            =   720
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "POSABS Y2"
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtOffsetY2 
         Alignment       =   1  'Right Justify
         DataField       =   "Y OFFSET 2"
         DataSource      =   "Data1"
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
         Left            =   2280
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "POSABS X2"
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Y"
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
         Left            =   1680
         TabIndex        =   98
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "X"
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
         Left            =   120
         TabIndex        =   96
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   93
      Text            =   "XXXX"
      Top             =   360
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12360
      Top             =   720
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3030
      Left            =   3720
      Picture         =   "118 Test.frx":0CCA
      Top             =   3000
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label LabelDBMode 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DB Mode"
      Height          =   300
      Left            =   11640
      TabIndex        =   187
      Top             =   10080
      Width           =   1035
   End
   Begin VB.Label LabelFC 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   6360
      TabIndex        =   168
      Top             =   11400
      Width           =   615
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
      Height          =   300
      Left            =   10680
      TabIndex        =   166
      Top             =   10920
      Width           =   1995
   End
   Begin VB.Label lblDate2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      Height          =   300
      Left            =   10680
      TabIndex        =   165
      Top             =   11280
      Width           =   1995
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default: "
      Height          =   300
      Left            =   10680
      TabIndex        =   164
      Top             =   10560
      Width           =   1995
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   10680
      TabIndex        =   163
      Top             =   11640
      Width           =   1995
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   6360
      TabIndex        =   65
      Top             =   10920
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SetObjCharString,SetObjSize,SetObjProfile"
      Height          =   255
      Left            =   2040
      TabIndex        =   146
      Top             =   2400
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Label lblInfo 
      Caption         =   "Position 2 :"
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
      Left            =   8040
      TabIndex        =   157
      Top             =   10920
      Width           =   1065
   End
   Begin VB.Label lblInfo 
      Caption         =   "Position 1 :"
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
      Left            =   8040
      TabIndex        =   156
      Top             =   10560
      Width           =   1035
   End
   Begin VB.Label LabelMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4440
      TabIndex        =   155
      Top             =   9960
      Width           =   2460
   End
   Begin VB.Label LabelPause 
      Caption         =   "Error Correction : "
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
      TabIndex        =   154
      Top             =   12000
      Width           =   1935
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      DataField       =   "MATRIX ID"
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
      Left            =   7200
      TabIndex        =   151
      Top             =   3720
      Width           =   555
   End
   Begin VB.Label Label13 
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
      TabIndex        =   150
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label Label12 
      Caption         =   "Case_ID:"
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
      Left            =   9600
      TabIndex        =   149
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [TBL CONFIG]"
      Height          =   300
      Left            =   14280
      TabIndex        =   136
      Top             =   12240
      Width           =   2715
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM [FIXTURE]"
      Height          =   300
      Left            =   7680
      TabIndex        =   135
      Top             =   9240
      Width           =   2715
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "timer1 I/O Status"
      Height          =   300
      Left            =   10680
      TabIndex        =   134
      Top             =   360
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Picture         =   "118 Test.frx":1A56C
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmTestScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub cmdFireDPPS_Click()

cmdFireDPPS.BackColor = &HFF&

X_OFFSET_1 = Val(txtOffsetX.Text)
Y_OFFSET_1 = Val(txtOffsetY.Text)
X_OFFSET_2 = Val(txtOffsetX2.Text)
Y_OFFSET_2 = Val(txtOffsetY2.Text)

If (optMarkRight.value = vbTrue) Then
        X_OFFSET = X_OFFSET_1
        Y_OFFSET = Y_OFFSET_1
End If

If (optMarkLeft.value = vbTrue) Then
        X_OFFSET = X_OFFSET_2
        Y_OFFSET = Y_OFFSET_2
End If
   
If (optLogo1.value = True) Then
    LOGO_MODE = LOGO_SIDE
    X_SHIFT = 0
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
    
    
If (OptionBoth.value = True) Then
    X_OFFSET = X_OFFSET_1
    Y_OFFSET = Y_OFFSET_1
    
   ' LaserPart
    
    X_OFFSET = X_OFFSET_2
    Y_OFFSET = Y_OFFSET_2
    
    'LaserPart
    
Else
    'LaserPart
End If
                  
'MsgBox "Complete", vbInformation, "Laser"

cmdFireDPPS.BackColor = &HFFC0C0

End Sub


Private Sub cmdIndex_Click()
'INDEX_COUNT = CLng(lblIndex.Caption)
cmdMovePocket_Click
End Sub

Private Sub cmdIndex2_Click()
'INDEX_COUNT = CLng(lblIndex2.Caption)
cmdMovePocket_Click
End Sub

Private Sub cmdLoad_Click()

Screen.MousePointer = vbHourglass

Dim sFilename As String
sFilename = "C:\MARKER\JOB\ATC DPSS.WLJ"

Dim JobIndex As Long

AutomationInterface.LoadJobFromFile sFilename, JobIndex
AutomationInterface.GetObjCount ObjectCount
 
Screen.MousePointer = vbDefault

'MsgBox "Load Job Complete", vbInformation, "Laser"

cmdLoad.BackColor = &H8000000F
cmdLoad.Enabled = False

End Sub

Private Sub cmdLoadParameters_Click()
cmdLoadParameters.Font.Bold = True
 
ACCEL_ID = CDbl(txtAccel.Text)
DECEL_ID = CDbl(txtDecel.Text)
VELOCITY_ID = CDbl(txtVel.Text)
    
'SetupRead

'Motor_Load_Parameters
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

Private Sub cmdPower_Click()
frmPowerFactors.Show vbModal
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


Private Sub cmdMove_Click()

'INDEX_COUNT = INDEX_PER_PULSE
cmdMovePocket_Click

End Sub

Private Sub cmdMovePocket_Click()

Dim axis
Dim moveComplete
Dim Error
Dim currentPos
Dim csr
Dim commandID
Dim resourceID
Dim errorCode

On Error GoTo Errorhandler
    
    'Set the boardID,axis and targetPosition

    axis = AXIS_ID
    
    moveComplete = 0
           
    Do
        'Load a target position
        Error = flex_load_target_pos(BOARD_ID, axis, INDEX_COUNT, &HFF)
        CheckError (Error)
                
        Error = flex_start(BOARD_ID, axis, 0)
        CheckError (Error)
            
        Do
            
            'Read the position
            Error = flex_read_pos_rtn(BOARD_ID, axis, currentPos)
            CheckError (Error)
        
            Error = flex_check_move_complete_status(BOARD_ID, axis, 0, moveComplete)
            CheckError (Error)

            'Delay 100 ms
            Sleep (100)
            
            'Check the modal errors
            flex_read_csr_rtn BOARD_ID, csr
            If (csr And NIMC_MODAL_ERROR_MSG) Then
            
                flex_stop_motion BOARD_ID, axis, NIMC_DECEL_STOP, 0 'Stop the Motion
                flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
                CheckError (errorCode)
            End If
            DoEvents 'allow other events to occur
        
        Loop Until (moveComplete)
        
    Loop Until moveComplete
    
    
    Exit Sub    'Exit the Sub
'////////////////////////////////////////////////////////////////////////
' Error Handling
'
Errorhandler:
    'First check for modal errors
    flex_read_csr_rtn BOARD_ID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn BOARD_ID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If
    
End Sub

Private Sub cmdOut_Click(Index As Integer)

If (OUTPUT_Status(Index - 1) = 0) Then
    OutputPort Index - 1, 1
    lblOut(Index).BackColor = &HC0FFC0
    Select Case Index
    Case 2
            lblOut(Index).Caption = "A"
    Case Else
            lblOut(Index).Caption = "On"
    End Select
    
Else
    OutputPort Index - 1, 0
    lblOut(Index).BackColor = &HC0C0FF
    
    Select Case Index
    Case 2
            lblOut(Index).Caption = "B"
    Case Else
            lblOut(Index).Caption = "Off"
     End Select
End If

SetOutputs frmTestScreen, 0

End Sub

Private Sub cmdProductionClear_Click()

Check1.value = vbUnchecked
Check2.value = vbChecked
Check3.value = vbChecked
Check4.value = vbChecked
Check5.value = vbChecked
Check6.value = vbChecked
Check7.value = vbChecked
Check8.value = vbUnchecked

Select Case AXIS_ID
Case 1
        OutputPort 2, 0
        lblOut(3).Caption = "Off"
        lblOut(3).BackColor = &HC0C0FF
Case 2
        OutputPort 3, 0
        lblOut(4).Caption = "Off"
        lblOut(4).BackColor = &HC0C0FF
End Select

cmdProductionRun_Click

Select Case AXIS_ID
Case 1
        OutputPort 2, 0
        lblOut(3).Caption = "Off"
        lblOut(3).BackColor = &HC0C0FF
Case 2
        OutputPort 3, 0
        lblOut(4).Caption = "Off"
        lblOut(4).BackColor = &HC0C0FF
End Select

End Sub

Private Sub cmdProductionRun_Click()

Dim ERROR1 As Integer
Dim ERROR2 As Integer
Dim PARTPRESENT As Integer

Dim delay As Double
delay = Val(txtPause.Text)

StartTime = Timer

LabelMessage.Caption = "Run Mode"

'MotionStartWR_INDEX_PER_PULSE

EXIT_INIT_ID = 0

Do
    If (EXIT_INIT_ID > 0) Then
        EXIT_INIT_ID = 0
        Exit Do
    End If
    
  '  MotionStatus (FREAD)
    DoEvents
        
     If (EXIT_ID > 0) Then
        
        Pause (delay)
                               
        If (Check9.value = vbChecked) Then
        
            If (InputPort(INPUT_6A) = 1) Then
                    MsgBox "Interlock Error", vbInformation, "ATC DPSS Laser"
                    Exit Do
            End If
            
        End If
        
        If (Check2.value = vbChecked) Then
                Select Case AXIS_ID
                Case 1
                        ERROR1 = InputPort(INPUT_6A)
                Case 2
                        ERROR1 = InputPort(INPUT_3A)
                End Select
                If (ERROR1 = 0) Then
                   MsgBox "Dial Error 1", vbCritical + vbInformation, "ATC DPSS Laser"
                   Exit Do
                End If
        End If
        
         If (Check3.value = vbChecked) Then
                Select Case AXIS_ID
                Case 1      'A SIDE
                        ERROR2 = InputPort(INPUT_7A)
                Case 2      'B SIDE
                        ERROR2 = InputPort(INPUT_4A)
                End Select
        
                If (ERROR2 = 0) Then
                   MsgBox "Dial Error 2", vbCritical + vbInformation, "ATC DPSS Laser"
                   Exit Do
                End If
        
        End If
        
        PARTPRESENT = 1
        LabelMessage.Caption = "Wait PP"
        Do Until (PARTPRESENT = 0)
                            
                 If (Check1.value = vbChecked) Then
                        Select Case AXIS_ID
                        Case 1
                                PARTPRESENT = InputPort(INPUT_5A)
                        Case 2
                                PARTPRESENT = InputPort(INPUT_2A)
                        End Select
                Else
                                PARTPRESENT = 0
                End If
                
                If (EXIT_INIT_ID > 0) Then
                   ' MotionStop
                    EXIT_INIT_ID = 1
                    Exit Do
                End If
                
                DoEvents
        Loop
        
        LabelMessage.Caption = "Run Mode"

        If (EXIT_INIT_ID = 1) Then
            Exit Do
        End If
                         
      '  Move_Motor
                
        DoEvents
    
        '
        'LASER POSITON RIGHT/LEFT CORRECTION OF A SIDE
        '
        Select Case HANDLER_ID
        Case 1  ' A SIDE
                If (Check4.value = vbChecked) Then
                    X_OFFSET = X_OFFSET_2
                    Y_OFFSET = Y_OFFSET_2
                    If (Check6.value = vbChecked) Then
                        If (COUNT_ID < POSITION_2 - 1) Then
                            'DO NOT LASER PART
                        Else
                            'LASER PART
                            'LaserPart
                            FIRE_COUNT_ID = FIRE_COUNT_ID + 1
                            LabelFC.Caption = FIRE_COUNT_ID
                        End If
                    Else
                       ' LaserPart
                    End If
                End If
                
                If (Check5.value = vbChecked) Then
                    X_OFFSET = X_OFFSET_1
                    Y_OFFSET = Y_OFFSET_1
                    If (Check7.value = vbChecked) Then
                        If (COUNT_ID < POSITION_1 - 1) Then
                            'DO NOT LASER PART
                        Else
                            'LASER PART
                            'LaserPart
                        End If
                    Else
                        'LaserPart
                    End If
                End If
        Case 2 'B SIDE
                If (Check4.value = vbChecked) Then
                    X_OFFSET = X_OFFSET_1
                    Y_OFFSET = Y_OFFSET_1
                    If (Check7.value = vbChecked) Then
                        If (COUNT_ID < POSITION_2 - 1) Then
                            'DO NOT LASER PART
                        Else
                            'LASER PART
                            'LaserPart
                        End If
                    Else
                       ' LaserPart
                    End If
                End If
                
                If (Check5.value = vbChecked) Then
                    X_OFFSET = X_OFFSET_2
                    Y_OFFSET = Y_OFFSET_2
                    If (Check6.value = vbChecked) Then
                        If (COUNT_ID < POSITION_1 - 1) Then
                            'DO NOT LASER PART
                        Else
                            'LASER PART
                           ' LaserPart
                            FIRE_COUNT_ID = FIRE_COUNT_ID + 1
                            LabelFC.Caption = FIRE_COUNT_ID
                        End If
                    Else
                       ' LaserPart
                    End If
                End If
        End Select
        
        POSITION_ID = POSITION_ID + 1
        COUNT_ID = COUNT_ID + 1
        lblPosition.Caption = POSITION_ID
        TextCOUNT.Text = COUNT_ID
        
        If (POSITION_ID = POCKETS_PER_REV) Then
            POSITION_ID = 0
        End If
        '
        'DIAL STOPS FOR FIRST PIECE INSPECTION
        '
        Select Case HANDLER_ID
        Case 1  ' A SIDE
                    If (Check7.value = vbChecked) Then
                            If (COUNT_ID = POSITION_1) Then
                                    MsgBox "Dial Stop 1", vbInformation, "ATC DPSS Laser"
                                 '   MotionStop
                                    EXIT_INIT_ID = 1
                                    Exit Do
                            End If
                    End If
                    If (Check6.value = vbChecked) Then
                            If (COUNT_ID = POSITION_2) Then
                                    MsgBox "Dial Stop 2", vbInformation, "ATC DPSS Laser"
                                 '   MotionStop
                                    EXIT_INIT_ID = 1
                                    Exit Do
                            End If
                    End If
        Case 2      'B CASE
                    If (Check6.value = vbChecked) Then
                            If (COUNT_ID = POSITION_1) Then
                                    MsgBox "Dial Stop 1", vbInformation, "ATC DPSS Laser"
                                '    MotionStop
                                    EXIT_INIT_ID = 1
                                    Exit Do
                            End If
                    End If
                    If (Check7.value = vbChecked) Then
                            If (COUNT_ID = POSITION_2) Then
                                    MsgBox "Dial Stop 2", vbInformation, "ATC DPSS Laser"
                                 '   MotionStop
                                    EXIT_INIT_ID = 1
                                    Exit Do
                            End If
                    End If
        End Select
        
        If (Check8.value = vbChecked) Then
                  '  MotionStop
                    EXIT_INIT_ID = 1
                    Exit Do
        End If
        
        If (Check9.value = vbChecked) Then
        
            If (InputPort(INPUT_6A) = 1) Then
                    'MotionStop
                    EXIT_INIT_ID = 1
                    MsgBox "Interlock Error", vbInformation, "ATC DPSS Laser"
                    Exit Do
            End If
            
        End If
        
        
    End If
Loop

LabelMessage.Caption = "Exit Run"

End Sub

Private Sub cmdProfile_Click()
frmProfile.Show
End Sub

Private Sub cmdResetCount_Click()
COUNT_ID = 0
TextCOUNT.Text = 0
End Sub

Private Sub cmdResetPosition_Click()
POSITION_ID = 0
lblPosition.Caption = POSITION_ID
End Sub

Private Sub cmdSetObj_Click()

'=============================================
'   LOGO MODE
'=============================================
X_OFFSET_1 = Val(txtOffsetX.Text)
Y_OFFSET_1 = Val(txtOffsetY.Text)
X_OFFSET_2 = Val(txtOffsetX2.Text)
Y_OFFSET_2 = Val(txtOffsetY2.Text)

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

Dim H_WIDTH(4) As Double
Dim sBuff As String

sBuff = Trim(Text1.Text)

If (Len(sBuff & "X") = 1) Then
    MsgBox "No text to Mark", vbCritical, "DPSS Laser"
    Exit Sub
End If

H_WIDTH(1) = Len(sBuff)
Select Case H_WIDTH(1)
Case 1
        H_WIDTH(1) = 0.25
Case 2
        H_WIDTH(1) = 0.5
Case 3
        H_WIDTH(1) = 0.75
Case Else
        H_WIDTH(1) = 1
End Select
sBuff = Trim(Text2.Text)
H_WIDTH(2) = Len(sBuff)
Select Case H_WIDTH(2)
Case 1
        H_WIDTH(2) = 0.25
Case 2
        H_WIDTH(2) = 0.5
Case 3
        H_WIDTH(2) = 0.75
Case Else
        H_WIDTH(2) = 1
End Select
sBuff = Trim(Text3.Text)
H_WIDTH(3) = Len(sBuff)
Select Case H_WIDTH(3)
Case 1
        H_WIDTH(3) = 0.25
Case 2
        H_WIDTH(3) = 0.5
Case 3
        H_WIDTH(3) = 0.75
Case Else
        H_WIDTH(3) = 1
End Select
sBuff = Trim(Text4.Text)
H_WIDTH(4) = Len(sBuff)
Select Case H_WIDTH(4)
Case 1
        H_WIDTH(4) = 0.25
Case 2
        H_WIDTH(4) = 0.5
Case 3
        H_WIDTH(4) = 0.75
Case Else
        H_WIDTH(4) = 1
End Select

'=============================================
'[1] SetObjCharString
'=============================================
Select Case MARK_MODE
Case 1
        sBuff = Trim(Text1.Text)
        AutomationInterface.SetObjCharString 0, sBuff
Case 2
        sBuff = Trim(Text1.Text)
        AutomationInterface.SetObjCharString 0, sBuff
        sBuff = Trim(Text2.Text)
        AutomationInterface.SetObjCharString 1, sBuff
Case 3
        sBuff = Trim(Text1.Text)
        AutomationInterface.SetObjCharString 0, sBuff
        sBuff = Trim(Text2.Text)
        AutomationInterface.SetObjCharString 1, sBuff
        sBuff = Trim(Text3.Text)
        AutomationInterface.SetObjCharString 2, sBuff
Case 4
        sBuff = Trim(Text1.Text)
        AutomationInterface.SetObjCharString 0, sBuff
        sBuff = Trim(Text2.Text)
        AutomationInterface.SetObjCharString 1, sBuff
        sBuff = Trim(Text3.Text)
        AutomationInterface.SetObjCharString 2, sBuff
        sBuff = Trim(Text4.Text)
        AutomationInterface.SetObjCharString 3, sBuff
End Select

'=============================================
'[2] SetObjSize
'=============================================

Dim sSQL As String
Set FR_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim HTextSize As Long
Dim VTextSize As Long

X_SHIFT = FR_Table.Fields("[X SHIFT]")

HTextSize = Format(FR_Table.Fields("[TEXT XSIZE]") * INCHES_TO_BITS, "0")
VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS * H_WIDTH(1), "0")
AutomationInterface.SetObjSize 0, HTextSize, VTextSize
VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS * H_WIDTH(2), "0")
AutomationInterface.SetObjSize 1, HTextSize, VTextSize
VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS * H_WIDTH(3), "0")
AutomationInterface.SetObjSize 2, HTextSize, VTextSize
VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS * H_WIDTH(4), "0")
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

'=============================================
'[3] SetObjProfile
'=============================================

Set FR_Database = OpenDatabase(ATC_LASER_BD)
sSQL = "SELECT * FROM [FIXTURE] WHERE [MATRIX ID] =" & MATRIX_ID
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

Dim i As Integer
For i = 0 To ObjectCount - 1
        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        AutomationInterface.SetObjProfile i, ProfileIndex, Markspeed_Bits, Jumpspeed_Bits, Jumpdelay, Markdelay, Polygondelay, Laserpower, Laseroffdelay, Laserondelay, Eightbitword, T1, T2, zAxis, Varijumpdelay, Varijumplength, Wobblesize, Wobblefrequency, Powerreset, Varipolydelay
Next i

sBuff = "SetObjProfiles"

'=======================================

Set FR_Database = OpenDatabase(ATC_LASER_BD)

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
Case Else
    ObjHPosition(4) = FR_Table.Fields("[LOGO LX TOP]")
    ObjVPosition(4) = FR_Table.Fields("[LOGO LY TOP]")
End Select
ObjHPosition(5) = FR_Table.Fields("[ATC  X]")
ObjVPosition(5) = FR_Table.Fields("[ATC  Y]")
FR_Table.Close
FR_Database.Close

'MsgBox "Setup Complete", vbInformation, "Laser"

cmdSetObj.Caption = "Setup Complete"

End Sub



Private Sub cmdStop_Click()
'MotionStop
EXIT_INIT_ID = 1
End Sub

Private Sub cmdUpdate_Click()

Data1.UpdateRecord

If (cmdLoad.Enabled = False) Then
    
'    cmdSetObj_Click

End If

End Sub

Private Sub CommandConfiguration_Click()
frmConfiguration.Show
End Sub

Private Sub CommandDN1_Click()
txtOffsetY2.Text = Format(Val(txtOffsetY2.Text) - 0.001, "0.000")
End Sub

Private Sub CommandDN11_Click()
txtOffsetY.Text = Format(Val(txtOffsetY.Text) - 0.001, "0.000")
End Sub

Private Sub CommandDN2_Click()
txtOffsetY2.Text = Format(Val(txtOffsetY2.Text) - 0.01, "0.000")
End Sub

Private Sub CommandDN22_Click()
txtOffsetY.Text = Format(Val(txtOffsetY.Text) - 0.01, "0.000")
End Sub



Private Sub CommandExit_Click()
Unload Me
End Sub

Private Sub CommandP_Click()
'RUN_MODE = 1
Form_Activate
End Sub

Private Sub CommandProductionRun_Click()

Check1.value = vbChecked
Check2.value = vbChecked
Check3.value = vbChecked
Check4.value = vbChecked
Check5.value = vbChecked
Check6.value = vbChecked
Check7.value = vbChecked
Check8.value = vbUnchecked

Select Case AXIS_ID
Case 1
        OutputPort 2, 1
        lblOut(3).Caption = "On"
        lblOut(3).BackColor = &HC0FFC0
Case 2
        OutputPort 3, 1
        lblOut(4).Caption = "On"
        lblOut(4).BackColor = &HC0FFC0
End Select

cmdProductionRun_Click

Select Case AXIS_ID
Case 1
        OutputPort 2, 0
        lblOut(3).Caption = "Off"
        lblOut(3).BackColor = &HC0C0FF
Case 2
        OutputPort 3, 0
        lblOut(4).Caption = "Off"
        lblOut(4).BackColor = &HC0C0FF
End Select

End Sub

Private Sub CommandProductionRun1_Click()

Check1.value = vbChecked
Check2.value = vbChecked
Check3.value = vbChecked
Check4.value = vbChecked
Check5.value = vbChecked
Check6.value = vbChecked
Check7.value = vbChecked
Check8.value = vbChecked

Select Case AXIS_ID
Case 1
        OutputPort 2, 1
        lblOut(3).Caption = "On"
        lblOut(3).BackColor = &HC0FFC0
Case 2
        OutputPort 3, 1
        lblOut(4).Caption = "On"
        lblOut(4).BackColor = &HC0FFC0
End Select

cmdProductionRun_Click

Select Case AXIS_ID
Case 1
        OutputPort 2, 0
        lblOut(3).Caption = "Off"
        lblOut(3).BackColor = &HC0C0FF
Case 2
        OutputPort 3, 0
        lblOut(4).Caption = "Off"
        lblOut(4).BackColor = &HC0C0FF
End Select

End Sub

Private Sub CommandPULSE1_Click()

If (lblIndex2.Caption > 10) Then
    lblIndex2.Caption = lblIndex2.Caption - 10
End If

End Sub

Private Sub CommandPULSE2_Click()
If (lblIndex2.Caption > 1) Then
    lblIndex2.Caption = lblIndex2.Caption - 1
End If
End Sub

Private Sub CommandPULSE3_Click()
If (lblIndex2.Caption < 50) Then
    lblIndex2.Caption = lblIndex2.Caption + 1
End If
End Sub

Private Sub CommandPULSE4_Click()
If (lblIndex2.Caption < 50) Then
    lblIndex2.Caption = lblIndex2.Caption + 10
End If
End Sub

Private Sub CommandPULSE5_Click()
If (lblIndex.Caption > 10) Then
    lblIndex.Caption = lblIndex.Caption - 10
End If
End Sub

Private Sub CommandPULSE6_Click()
If (lblIndex.Caption > 1) Then
    lblIndex.Caption = lblIndex.Caption - 1
End If
End Sub

Private Sub CommandPULSE7_Click()
If (lblIndex.Caption < INDEX_PER_PULSE) Then
    lblIndex.Caption = lblIndex.Caption + 1
End If
End Sub

Private Sub CommandPULSE8_Click()
If (lblIndex.Caption < INDEX_PER_PULSE) Then
    lblIndex.Caption = lblIndex.Caption + 10
End If
End Sub

Private Sub CommandResetFireCount_Click()

FIRE_COUNT_ID = 0
LabelFC.Caption = FIRE_COUNT_ID

End Sub

Private Sub CommandSet_Click()

COUNT_ID = Val(TextCOUNT.Text)
 
End Sub

Private Sub CommandT_Click()
'RUN_MODE = 0
Form_Activate
End Sub

Private Sub CommandUP1_Click()
txtOffsetY2.Text = Format(Val(txtOffsetY2.Text) + 0.001, "0.000")
End Sub

Private Sub CommandUP11_Click()
txtOffsetY.Text = Format(Val(txtOffsetY.Text) + 0.001, "0.000")
End Sub

Private Sub CommandUP2_Click()
txtOffsetY2.Text = Format(Val(txtOffsetY2.Text) + 0.01, "0.000")
End Sub

Private Sub CommandUP22_Click()
txtOffsetY.Text = Format(Val(txtOffsetY.Text) + 0.01, "0.000")
End Sub

Private Sub Form_Activate()

Select Case LOGO_MODE
Case LOGO_ATC
        optLogo4.value = True
Case LOGO_SIDE
        optLogo1.value = True
End Select

Dim sSQL As String
sSQL = "SELECT * FROM [FIXTURE] WHERE [MATRIX ID] =" & MATRIX_ID
                                   
Data1.RecordSource = sSQL
Data1.Refresh

optHand_Click (HANDLER_ID)

cmdIndex.Caption = "Pulse [1-" & INDEX_PER_PULSE & "]"
lblIndex.Caption = INDEX_PER_PULSE

'MAIN AIR ON
OutputPort 0, 1
lblOut(1).Caption = "On"
lblOut(1).BackColor = &HC0FFC0

If TEXT_ID <> "NA" Then
    Text1.Text = TEXT_ID
End If

Select Case RUN_MODE
Case 1  'PRODUCTION MODE
        CommandConfiguration.Visible = False
        cmdProfile.Visible = False
        cmdFixture.Visible = False
        cmdPower.Visible = False
        CommandExit.Visible = True
        
        FramePower.Enabled = True
        FrameOptions.Enabled = False
        FrameHandler.Enabled = False
        
        cmdLoad.Visible = False
        cmdProductionRun.Enabled = False
        cmdFireDPPS.Visible = False
        optMarkLeft.Visible = False
        optMarkRight.Visible = False
        OptionBoth.Visible = False
        
        CommandP.BackColor = &HFFFFC0
        CommandT.BackColor = &H8000000F
        
Case 0
        CommandConfiguration.Visible = True
        cmdProfile.Visible = True
        cmdFixture.Visible = True
        cmdPower.Visible = True
        CommandExit.Visible = False
        
        FramePower.Enabled = True
        FrameOptions.Enabled = True
        FrameHandler.Enabled = True
        
        cmdLoad.Enabled = True
        cmdProductionRun.Enabled = True
        cmdFireDPPS.Visible = True
        optMarkLeft.Visible = True
        optMarkRight.Visible = True
        OptionBoth.Visible = True
        
        CommandP.BackColor = &H8000000F
        CommandT.BackColor = &HFFFFC0
End Select

End Sub


Private Sub Form_Load()

Caption = "DPSS Laser Test Panel       " & ATC_DWG & "       " & ATC_VERSION
 
lblDate2.Caption = Date
lblUser.Caption = strComputerName
lblIP.Caption = IP_ADDRESS
lblTime.Caption = Format(Time, "HH:MM AM/PM")
 
TextLOCATION.Text = LOCATION_ID
 
Select Case DataBase_MODE
Case DATABASE_MODE_REM_JR
        LabelDBMode.Caption = "REM JR"
Case DATABASE_MODE_REM_NY
        LabelDBMode.Caption = "REM NY"
Case DATABASE_MODE_LCL
        LabelDBMode.Caption = "LCL"
Case DATABASE_MODE_FIL
        LabelDBMode.Caption = "FILE"
End Select
 
 
InputOutput frmTestScreen

Data1.DatabaseName = ATC_LASER_BD
Data2.DatabaseName = ATC_LASER_BD

Dim Index As Integer
For Index = 1 To 8
    Select Case Index
    Case 1, 2, 3, 4, 5, 6, 7, 8
        If (OUTPUT_Status(Index + 7) = 1) Then
            lblOut(Index).Caption = "On"
            lblOut(Index).BackColor = &HC0FFC0
        Else
            lblOut(Index).Caption = "Off"
            lblOut(Index).BackColor = &HC0C0FF
        End If
    End Select
Next Index

Select Case RUN_MODE
Case 1  'PRODUCTION
        If (LOAD_COMPLETE = 0) Then
            Init_PCI_IO
            cmdLoad_Click
            LOAD_COMPLETE = 1
        End If
Case 0
        Init_PCI_IO
        cmdLoad_Click
        LOAD_COMPLETE = 1
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

Select Case RUN_MODE
Case 1
        frmTestScreen.Hide
        frmOPScreen.Show
Case 0

        Dim iAns As Integer
        iAns = MsgBox("Exit Program", vbYesNo, "DPSS Laser Test Panel ")
        If (iAns = vbYes) Then
            End
        Else
            Cancel = 1
            frmTestScreen.Show
        End If
End Select

End Sub


Private Sub optHand_Click(Index As Integer)

HANDLER_ID = Index
  
optHand(Index).value = True
  
Dim sSQL As String
 
sSQL = "SELECT * FROM [TBL CONFIG] WHERE [ID]=" & HANDLER_ID

Data2.RecordSource = sSQL
Data2.Refresh

'SetupRead

'Motor_Load_Parameters

cmdIndex.Caption = "Pulse [1-" & INDEX_PER_PULSE & "]"
lblIndex.Caption = INDEX_PER_PULSE

'111 DPSS Laser 01/21/2014 A/B Air
Select Case AXIS_ID
Case 1  'A
        OutputPort 1, 0
        lblOut(2).Caption = "A"
Case 2
        OutputPort 1, 1
        lblOut(2).Caption = "B"
End Select

End Sub


Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub
Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
End Sub
Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3)
End Sub
Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4)
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub cmdMinus1_Click()

txtOffsetX.Text = Format(Val(txtOffsetX.Text) + 0.001, "0.000")
 
End Sub

Private Sub cmdMinus11_Click()

txtOffsetX.Text = Format(Val(txtOffsetX.Text) + 0.01, "0.000")
 
End Sub

Private Sub cmdMinus2_Click()
 
txtOffsetX2.Text = Format(Val(txtOffsetX2.Text) + 0.001, "0.000")
 
End Sub

Private Sub cmdMinus22_Click()

txtOffsetX2.Text = Format(Val(txtOffsetX2.Text) + 0.01, "0.000")
 
End Sub
Private Sub cmdPlus1_Click()

txtOffsetX.Text = Format(Val(txtOffsetX.Text) - 0.001, "0.000")
 
End Sub

Private Sub cmdPlus11_Click()

txtOffsetX.Text = Format(Val(txtOffsetX.Text) - 0.01, "0.000")
 
End Sub

Private Sub cmdPlus2_Click()

txtOffsetX2.Text = Format(Val(txtOffsetX2.Text) - 0.001, "0.000")
 
End Sub

Private Sub cmdPlus22_Click()

txtOffsetX2.Text = Format(Val(txtOffsetX2.Text) - 0.01, "0.000")
 
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sBuff As String

sBuff = UCase(txtPassword.Text)

If (Button = 2 And Shift = 1) Then
       
    Select Case sBuff
        Case "MIKE" & Mid(Format(Date, "ddd"), 1, 1), "ERIK" & Mid(Format(Date, "ddd"), 1, 1), "KELLY" & Mid(Format(Date, "ddd"), 1, 1)

            LabelPause.Visible = True
            txtPause.Visible = True
            fraPTMode.Enabled = True
    Case Else
             
    End Select
 
Else
    'CommandBackUpDB.Visible = False
    LabelPause.Visible = False
    txtPause.Visible = False
    fraPTMode.Enabled = False
End If

  txtPassword.Text = "XXXX"
  
End Sub

Private Sub Timer1_Timer()

iInput(1) = InputPort(INPUT_5A)
iInput(2) = InputPort(INPUT_6A)
iInput(3) = InputPort(INPUT_7A)
iInput(4) = InputPort(INPUT_2A)
iInput(5) = InputPort(INPUT_3A)
iInput(6) = InputPort(INPUT_4A)
iInput(7) = InputPort(INPUT_8A)
iInput(8) = InputPort(INPUT_1A)
       
ShowIO frmTestScreen

End Sub
