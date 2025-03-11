VERSION 5.00
Begin VB.Form frm414 
   Caption         =   "Tray 414 C Case"
   ClientHeight    =   11820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18720
   LinkTopic       =   "Form1"
   ScaleHeight     =   11820
   ScaleWidth      =   18720
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      DataField       =   "CAMERA POS"
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
      Index           =   0
      Left            =   13680
      TabIndex        =   314
      Text            =   "CAMERA POS"
      ToolTipText     =   "CAMERA POS"
      Top             =   7800
      Width           =   1100
   End
   Begin VB.CommandButton CommandUpdateGlobal 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Global PWR"
      Height          =   300
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   313
      ToolTipText     =   "[FIXTURE] WHERE [MATRIX ID]"
      Top             =   11160
      Width           =   1300
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
      Left            =   120
      TabIndex        =   297
      Top             =   8400
      Width           =   1935
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
         TabIndex        =   303
         Top             =   360
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
         TabIndex        =   302
         Top             =   735
         Value           =   -1  'True
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
         TabIndex        =   301
         Top             =   1110
         Width           =   900
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
         TabIndex        =   300
         Top             =   1485
         Width           =   900
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
         TabIndex        =   299
         Top             =   2235
         Width           =   1500
      End
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
         TabIndex        =   298
         Top             =   1860
         Width           =   1380
      End
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
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   296
      Top             =   7080
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
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   295
      Top             =   8280
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
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   294
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton CommandLoadDPSS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Load DPSS"
      Height          =   300
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   293
      Top             =   7680
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update Record"
      Height          =   300
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   283
      ToolTipText     =   "[FIXTURE] WHERE [MATRIX ID]"
      Top             =   11160
      Width           =   1485
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
      Left            =   14040
      TabIndex        =   282
      Top             =   10560
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
      Left            =   16425
      TabIndex        =   281
      Top             =   10560
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
      Left            =   15630
      TabIndex        =   280
      Top             =   10560
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
      Left            =   14835
      TabIndex        =   279
      Top             =   10560
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
      Left            =   14040
      TabIndex        =   278
      Top             =   9600
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
      Left            =   16440
      TabIndex        =   277
      Top             =   9600
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
      Left            =   15630
      TabIndex        =   276
      Top             =   9600
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
      Left            =   14835
      TabIndex        =   275
      Top             =   9600
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
      Left            =   14040
      TabIndex        =   274
      Top             =   10080
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
      Left            =   16425
      TabIndex        =   273
      Top             =   10080
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
      Left            =   15630
      TabIndex        =   272
      Top             =   10080
      Width           =   800
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
      Left            =   14835
      TabIndex        =   271
      Top             =   10080
      Width           =   800
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
      Left            =   12840
      TabIndex        =   270
      Text            =   "X"
      ToolTipText     =   "PulseWidth"
      Top             =   10560
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
      Left            =   12840
      TabIndex        =   269
      Text            =   "XXXXXX"
      Top             =   10080
      Width           =   1095
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
      Left            =   12840
      TabIndex        =   268
      Text            =   "XXX"
      Top             =   9600
      Width           =   1095
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
      Left            =   4680
      TabIndex        =   253
      Top             =   9000
      Width           =   3855
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
         Left            =   2760
         TabIndex        =   262
         Top             =   960
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
         Left            =   3240
         TabIndex        =   261
         Top             =   960
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
         Left            =   3000
         TabIndex        =   260
         Top             =   1440
         Width           =   375
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
         Left            =   3000
         TabIndex        =   259
         Top             =   480
         Width           =   375
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
         TabIndex        =   258
         Text            =   "30.0"
         ToolTipText     =   "L X Offset"
         Top             =   960
         Width           =   700
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
         Left            =   1800
         TabIndex        =   257
         Text            =   "30.0"
         ToolTipText     =   "L Y Offset"
         Top             =   960
         Width           =   700
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
         Left            =   1800
         TabIndex        =   256
         Text            =   "0"
         Top             =   480
         Width           =   700
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
         TabIndex        =   255
         Text            =   "0"
         Top             =   480
         Width           =   700
      End
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
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   254
         Text            =   "0"
         ToolTipText     =   "OFFSET"
         Top             =   1560
         Width           =   495
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
         Left            =   1440
         TabIndex        =   267
         Top             =   480
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
         TabIndex        =   266
         Top             =   480
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
         TabIndex        =   265
         Top             =   960
         Width           =   375
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
         Left            =   1440
         TabIndex        =   264
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
         Left            =   240
         TabIndex        =   263
         Top             =   1560
         Width           =   1365
      End
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   252
      Top             =   8040
      Width           =   1575
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
      Left            =   13320
      TabIndex        =   251
      Text            =   "0"
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   " Tray Configuration "
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
      Height          =   1695
      Left            =   15480
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton OptionTray2 
         Caption         =   "[2] 9 X 4 Chip"
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
         TabIndex        =   37
         Top             =   360
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton OptionTray3 
         Caption         =   "[3] 4 X 4 Wire"
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
         TabIndex        =   36
         Top             =   720
         Width           =   1620
      End
      Begin VB.OptionButton OptionTray4 
         Caption         =   "[4] 3 X 4 Lead"
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
         TabIndex        =   35
         Top             =   1080
         Width           =   1620
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 [TBL Size Location]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   3000
      Visible         =   0   'False
      Width           =   3420
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
      Left            =   15495
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   8280
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
      Height          =   375
      Left            =   15870
      Style           =   1  'Graphical
      TabIndex        =   248
      Top             =   8280
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
      Height          =   375
      Left            =   16245
      Style           =   1  'Graphical
      TabIndex        =   247
      Top             =   8280
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
      Height          =   375
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   246
      Top             =   8280
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
      Height          =   375
      Left            =   13680
      MaxLength       =   10
      TabIndex        =   245
      Text            =   "Z"
      ToolTipText     =   "Z HEIGHT"
      Top             =   8280
      Width           =   1215
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   244
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   2640
      Visible         =   0   'False
      Width           =   3300
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   243
      Top             =   7800
      Width           =   1815
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   242
      Top             =   8520
      Width           =   1620
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
      Left            =   12360
      TabIndex        =   241
      Text            =   "6000"
      Top             =   7200
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
      Left            =   360
      TabIndex        =   240
      Text            =   "0"
      Top             =   7200
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
      Left            =   3480
      TabIndex        =   239
      Text            =   "2000"
      Top             =   7200
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
      Left            =   6600
      TabIndex        =   238
      Text            =   "4000"
      Top             =   7200
      Width           =   735
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
      Left            =   9480
      TabIndex        =   237
      Text            =   "6000"
      Top             =   7200
      Width           =   735
   End
   Begin VB.Data Data1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Data1 [TBL Tray Config]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   2280
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.OptionButton OptionAll 
      Caption         =   " [All] Segments 1-5"
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
      Left            =   8760
      TabIndex        =   236
      Top             =   10800
      Width           =   1980
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
      Index           =   4
      Left            =   13560
      TabIndex        =   235
      TabStop         =   0   'False
      Top             =   6000
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
      Index           =   4
      Left            =   13560
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   6375
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
      Index           =   4
      Left            =   12360
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   233
      TabStop         =   0   'False
      Top             =   6000
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
      Index           =   4
      Left            =   12360
      TabIndex        =   232
      TabStop         =   0   'False
      Top             =   6375
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
      Index           =   4
      Left            =   12360
      TabIndex        =   231
      TabStop         =   0   'False
      Top             =   6750
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
      Index           =   4
      Left            =   13560
      TabIndex        =   230
      TabStop         =   0   'False
      Top             =   6750
      Width           =   1200
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   81
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   228
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   80
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   227
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   83
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   226
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   82
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   71
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   70
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   73
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   222
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   72
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   60
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   220
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   61
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   219
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   63
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   62
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   217
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   50
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   216
      Top             =   3360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   51
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   215
      Top             =   3360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   53
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   3360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   52
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   3360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   40
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   212
      Top             =   2760
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   41
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   211
      Top             =   2760
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   43
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   210
      Top             =   2760
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   42
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   209
      Top             =   2760
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   30
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   208
      Top             =   2160
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   31
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   207
      Top             =   2160
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   33
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   206
      Top             =   2160
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   32
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   205
      Top             =   2160
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   20
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   1560
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   21
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   203
      Top             =   1560
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   23
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   1560
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   22
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   1560
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   10
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   200
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   11
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   199
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   13
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   198
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   12
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   197
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   2
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   3
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   1
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix4 
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
      Height          =   500
      Index           =   0
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "83"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   83
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   82
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   81
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   190
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   80
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "73"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   73
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   72
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   187
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   71
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   70
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "63"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   63
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   62
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   61
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   60
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   80
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   180
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   81
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   82
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "83"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   83
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   70
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   71
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   72
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   174
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "73"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   73
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   173
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   60
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   61
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   62
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "63"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   63
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "83"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   83
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   168
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   82
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   81
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   80
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "73"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   73
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   72
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   71
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   70
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "63"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   63
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   62
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   61
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   60
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   80
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   5280
      Width           =   500
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
      Height          =   500
      Index           =   81
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   5280
      Width           =   500
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
      Height          =   500
      Index           =   82
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "83"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   83
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   5280
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   70
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   4665
      Width           =   500
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
      Height          =   500
      Index           =   71
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   4680
      Width           =   500
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
      Height          =   500
      Index           =   72
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "73"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   73
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   4680
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   60
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   4050
      Width           =   500
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
      Height          =   500
      Index           =   61
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   4080
      Width           =   500
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
      Height          =   500
      Index           =   62
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   4080
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "63"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   63
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   4080
      Width           =   500
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
      Index           =   3
      Left            =   10680
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   6750
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
      Index           =   3
      Left            =   9480
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   6750
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
      Index           =   2
      Left            =   7800
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   6750
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
      Index           =   2
      Left            =   6600
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   6750
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
      Index           =   1
      Left            =   4680
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   6750
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
      Index           =   1
      Left            =   3480
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   6750
      Width           =   1200
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   30
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   0
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   1
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   3
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   10
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   11
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   12
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   13
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   20
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   21
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   22
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   23
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   31
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   33
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   40
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   41
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   42
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "43"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   43
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   50
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   51
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   52
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   53
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   32
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix3 
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
      Height          =   500
      Index           =   2
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   360
      Width           =   500
   End
   Begin VB.Frame Frame3 
      Caption         =   "Segment 4 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   9240
      TabIndex        =   20
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   53
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   52
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   51
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   50
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   43
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   42
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   41
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   40
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   33
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   32
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   31
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   23
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   22
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   21
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   20
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   13
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   12
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   11
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   10
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   3
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   2
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   1
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
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
      Height          =   500
      Index           =   0
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix2 
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   30
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   0
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   1
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   3
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   10
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   11
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   12
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   13
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   990
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   20
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   21
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   22
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   23
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   1605
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   30
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   31
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   32
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   33
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   2235
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   40
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   41
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   42
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   43
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   2850
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   50
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   51
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
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
      Height          =   500
      Index           =   52
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix1 
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   53
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   53
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3480
      Width           =   500
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
      Height          =   500
      Index           =   52
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3480
      Width           =   500
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
      Height          =   500
      Index           =   51
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   50
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3435
      Width           =   500
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
      Height          =   500
      Index           =   43
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2820
      Width           =   500
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
      Height          =   500
      Index           =   42
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2820
      Width           =   500
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
      Height          =   500
      Index           =   41
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2820
      Width           =   500
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
      Height          =   500
      Index           =   40
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2820
      Width           =   500
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
      Height          =   500
      Index           =   33
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2205
      Width           =   500
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
      Height          =   500
      Index           =   32
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2205
      Width           =   500
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
      Height          =   500
      Index           =   31
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2205
      Width           =   500
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
      Height          =   500
      Index           =   30
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2205
      Width           =   500
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
      Height          =   500
      Index           =   23
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1590
      Width           =   500
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
      Height          =   500
      Index           =   22
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1590
      Width           =   500
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
      Height          =   500
      Index           =   21
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1590
      Width           =   500
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
      Height          =   500
      Index           =   20
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   1590
      Width           =   500
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
      Height          =   500
      Index           =   13
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   960
      Width           =   500
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
      Height          =   500
      Index           =   12
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   960
      Width           =   500
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
      Height          =   500
      Index           =   11
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   960
      Width           =   500
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
      Height          =   500
      Index           =   10
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   975
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
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
      Height          =   500
      Index           =   3
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   360
      Width           =   500
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
      Height          =   500
      Index           =   2
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   360
      Width           =   500
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
      Height          =   500
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   360
      Width           =   500
   End
   Begin VB.CommandButton cmdMatrix 
      BackColor       =   &H00808080&
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
      Height          =   500
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   360
      Width           =   500
   End
   Begin VB.OptionButton Option5 
      Caption         =   "[5] Segments 5"
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
      Left            =   8760
      TabIndex        =   42
      Top             =   10440
      Width           =   1980
   End
   Begin VB.OptionButton Option4 
      Caption         =   "[4] Segment 4"
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
      Left            =   8760
      TabIndex        =   41
      Top             =   10080
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
      Height          =   360
      Left            =   8760
      TabIndex        =   40
      Top             =   9720
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
      Height          =   360
      Left            =   8760
      TabIndex        =   39
      Top             =   9360
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
      Height          =   360
      Left            =   8760
      TabIndex        =   38
      Top             =   9000
      Value           =   -1  'True
      Width           =   1740
   End
   Begin VB.CommandButton cmdTrayConfig 
      Caption         =   "Tray Config"
      Height          =   300
      Left            =   15960
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   1140
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
      Index           =   3
      Left            =   9480
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6375
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
      Index           =   3
      Left            =   9480
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6000
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
      Index           =   3
      Left            =   10680
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1200
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
      Index           =   3
      Left            =   10680
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6000
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
      Index           =   2
      Left            =   6600
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6375
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
      Index           =   2
      Left            =   6600
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6000
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
      Index           =   2
      Left            =   7800
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1200
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
      Index           =   2
      Left            =   7800
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6000
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
      Index           =   1
      Left            =   3480
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6375
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
      Index           =   1
      Left            =   3480
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6000
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
      Index           =   1
      Left            =   4680
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1200
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
      Index           =   1
      Left            =   4680
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Segment 3 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   6240
      TabIndex        =   19
      Top             =   0
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Segment 2 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   3240
      TabIndex        =   18
      Top             =   0
      Width           =   2895
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
      Left            =   8880
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   11280
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
      Left            =   2160
      TabIndex        =   8
      Top             =   9000
      Width           =   2415
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
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   1245
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
         Left            =   960
         TabIndex        =   16
         Top             =   855
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
         Left            =   960
         TabIndex        =   15
         Top             =   1230
         Visible         =   0   'False
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
         Left            =   960
         TabIndex        =   14
         Top             =   1605
         Visible         =   0   'False
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
         TabIndex        =   13
         Top             =   480
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
         TabIndex        =   12
         Top             =   855
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
         TabIndex        =   11
         Top             =   1230
         Visible         =   0   'False
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
         TabIndex        =   10
         Top             =   1605
         Visible         =   0   'False
         Width           =   645
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
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6000
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
      Left            =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6375
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
      Left            =   360
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6000
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
      Left            =   360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6375
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
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6750
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
      Left            =   1560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6750
      Width           =   1200
   End
   Begin VB.Frame fraM 
      Caption         =   "Segment 1 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Frame Frame5 
      Caption         =   "Segment 5 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   12240
      TabIndex        =   229
      Top             =   0
      Width           =   2895
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
      Left            =   15960
      TabIndex        =   312
      Top             =   4920
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
      Left            =   17520
      TabIndex        =   311
      Top             =   4920
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
      Left            =   15960
      TabIndex        =   310
      Top             =   4440
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
      Left            =   17520
      TabIndex        =   309
      Top             =   4440
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
      Left            =   15960
      TabIndex        =   308
      Top             =   3960
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
      Left            =   17520
      TabIndex        =   307
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      DataField       =   "ROTATION"
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
      Left            =   17520
      TabIndex        =   306
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "ROTATION_ID:"
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
      Left            =   15960
      TabIndex        =   305
      Top             =   5400
      Width           =   1335
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
      Left            =   17520
      TabIndex        =   304
      Top             =   1920
      Width           =   885
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
      Left            =   15360
      TabIndex        =   292
      Top             =   8880
      Width           =   1845
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
      Left            =   14400
      TabIndex        =   291
      Top             =   8880
      Width           =   765
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
      Left            =   12600
      TabIndex        =   290
      Top             =   8880
      Width           =   1485
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "[2 to 65535]"
      Height          =   195
      Left            =   11400
      TabIndex        =   289
      Top             =   10800
      Width           =   855
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "[0 to 30000]"
      Height          =   195
      Left            =   11400
      TabIndex        =   288
      Top             =   10320
      Width           =   855
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "[0.02 to 250.0]"
      Height          =   195
      Left            =   11400
      TabIndex        =   287
      Top             =   9840
      Width           =   1035
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
      Left            =   10920
      TabIndex        =   286
      Top             =   10560
      Width           =   1440
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
      Left            =   10920
      TabIndex        =   285
      Top             =   10080
      Width           =   1755
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
      Left            =   10920
      TabIndex        =   284
      Top             =   9600
      Width           =   1470
   End
   Begin VB.Label lblSegment 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEGMENT #"
      Height          =   375
      Left            =   4680
      TabIndex        =   250
      Top             =   7800
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
      Left            =   240
      TabIndex        =   1
      Top             =   7800
      Width           =   2295
   End
End
Attribute VB_Name = "frm414"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAll_Click(Index As Integer)

Dim X As Integer, Y As Integer, k As Integer
For X = 0 To 3
       For Y = 0 To 8
            k = X + (10 * Y)
            Select Case Index
            Case 0
                    If (cmdMatrix(k).Enabled = True) Then
                        cmdMatrix(k).Caption = k
                        cmdMatrix(k).BackColor = &HC0FFFF
                    End If
            Case 1
                    If (cmdMatrix1(k).Enabled = True) Then
                        cmdMatrix1(k).Caption = k
                        cmdMatrix1(k).BackColor = &HC0FFFF
                    End If
            Case 2
                    If (cmdMatrix2(k).Enabled = True) Then
                        cmdMatrix2(k).Caption = k
                        cmdMatrix2(k).BackColor = &HC0FFFF
                    End If
            Case 3
                    If (cmdMatrix3(k).Enabled = True) Then
                        cmdMatrix3(k).Caption = k
                        cmdMatrix3(k).BackColor = &HC0FFFF
                    End If
            Case 4
                    If (cmdMatrix4(k).Enabled = True) Then
                        cmdMatrix4(k).Caption = k
                        cmdMatrix4(k).BackColor = &HC0FFFF
                    End If
            End Select
            
        Next Y
Next X


End Sub


Private Sub cmdColumn_Click(Index As Integer)

Dim Y As Integer
Dim i As Integer
 
For Y = 0 To 3
     Select Case Index
     Case 0
                If (cmdMatrix(Y).Caption <> "") Then
                    For i = 0 To 8
                        If cmdMatrix(i * 10 + Y).Enabled = True Then
                            cmdMatrix(i * 10 + Y).Caption = i * 10 + Y
                            cmdMatrix(i * 10 + Y).BackColor = &HC0FFFF
                        End If
                    Next i
                End If
     Case 1
                If (cmdMatrix1(Y).Caption <> "") Then
                    For i = 0 To 8
                        If cmdMatrix1(i * 10 + Y).Enabled = True Then
                           cmdMatrix1(i * 10 + Y).Caption = i * 10 + Y
                           cmdMatrix1(i * 10 + Y).BackColor = &HC0FFFF
                        End If
                    Next i
                End If
     Case 2
                If (cmdMatrix2(Y).Caption <> "") Then
                    For i = 0 To 8
                        If cmdMatrix2(i * 10 + Y).Enabled = True Then
                           cmdMatrix2(i * 10 + Y).Caption = i * 10 + Y
                           cmdMatrix2(i * 10 + Y).BackColor = &HC0FFFF
                        End If
                    Next i
                End If
     Case 3
                If (cmdMatrix3(Y).Caption <> "") Then
                    For i = 0 To 8
                        If cmdMatrix3(i * 10 + Y).Enabled = True Then
                           cmdMatrix3(i * 10 + Y).Caption = i * 10 + Y
                           cmdMatrix3(i * 10 + Y).BackColor = &HC0FFFF
                        End If
                    Next i
                End If
     Case 4
                If (cmdMatrix4(Y).Caption <> "") Then
                    For i = 0 To 8
                        If cmdMatrix4(i * 10 + Y).Enabled = True Then
                           cmdMatrix4(i * 10 + Y).Caption = i * 10 + Y
                           cmdMatrix4(i * 10 + Y).BackColor = &HC0FFFF
                        End If
                    Next i
                End If
     End Select
Next Y

End Sub

Private Sub cmdCorners_Click(Index As Integer)


Select Case TRAY_ID
Case 4, 12
        Select Case Index
        Case 0
                cmdMatrix(10).Caption = 10
                cmdMatrix(13).Caption = 13
                cmdMatrix(70).Caption = 70
                cmdMatrix(73).Caption = 73
                cmdMatrix(10).BackColor = &HC0FFFF
                cmdMatrix(13).BackColor = &HC0FFFF
                cmdMatrix(70).BackColor = &HC0FFFF
                cmdMatrix(73).BackColor = &HC0FFFF
                                
        Case 1
                cmdMatrix1(10).Caption = 10
                cmdMatrix1(13).Caption = 13
                cmdMatrix1(70).Caption = 70
                cmdMatrix1(73).Caption = 73
        
                cmdMatrix1(10).BackColor = &HC0FFFF
                cmdMatrix1(13).BackColor = &HC0FFFF
                cmdMatrix1(70).BackColor = &HC0FFFF
                cmdMatrix1(73).BackColor = &HC0FFFF
        Case 2
                cmdMatrix2(10).Caption = 10
                cmdMatrix2(13).Caption = 13
                cmdMatrix2(70).Caption = 70
                cmdMatrix2(73).Caption = 73
                cmdMatrix2(10).BackColor = &HC0FFFF
                cmdMatrix2(13).BackColor = &HC0FFFF
                cmdMatrix2(70).BackColor = &HC0FFFF
                cmdMatrix2(73).BackColor = &HC0FFFF
        Case 3
                cmdMatrix3(10).Caption = 10
                cmdMatrix3(13).Caption = 13
                cmdMatrix3(70).Caption = 70
                cmdMatrix3(73).Caption = 73
                cmdMatrix3(10).BackColor = &HC0FFFF
                cmdMatrix3(13).BackColor = &HC0FFFF
                cmdMatrix3(70).BackColor = &HC0FFFF
                cmdMatrix3(73).BackColor = &HC0FFFF
        Case 4
                cmdMatrix4(10).Caption = 10
                cmdMatrix4(13).Caption = 13
                cmdMatrix4(70).Caption = 70
                cmdMatrix4(73).Caption = 73
                cmdMatrix4(10).BackColor = &HC0FFFF
                cmdMatrix4(13).BackColor = &HC0FFFF
                cmdMatrix4(70).BackColor = &HC0FFFF
                cmdMatrix4(73).BackColor = &HC0FFFF
        End Select

Case Else
        Select Case Index
        Case 0
                cmdMatrix(0).Caption = 0
                cmdMatrix(3).Caption = 3
                cmdMatrix(80).Caption = 80
                cmdMatrix(83).Caption = 83
                cmdMatrix(0).BackColor = &HC0FFFF
                cmdMatrix(3).BackColor = &HC0FFFF
                cmdMatrix(80).BackColor = &HC0FFFF
                cmdMatrix(83).BackColor = &HC0FFFF
                                
        Case 1
                cmdMatrix1(0).Caption = 0
                cmdMatrix1(3).Caption = 3
                cmdMatrix1(80).Caption = 80
                cmdMatrix1(83).Caption = 83
        
                cmdMatrix1(0).BackColor = &HC0FFFF
                cmdMatrix1(3).BackColor = &HC0FFFF
                cmdMatrix1(80).BackColor = &HC0FFFF
                cmdMatrix1(83).BackColor = &HC0FFFF
        Case 2
                cmdMatrix2(0).Caption = 0
                cmdMatrix2(3).Caption = 3
                cmdMatrix2(80).Caption = 80
                cmdMatrix2(83).Caption = 83
                cmdMatrix2(0).BackColor = &HC0FFFF
                cmdMatrix2(3).BackColor = &HC0FFFF
                cmdMatrix2(80).BackColor = &HC0FFFF
                cmdMatrix2(83).BackColor = &HC0FFFF
        Case 3
                cmdMatrix3(0).Caption = 0
                cmdMatrix3(3).Caption = 3
                cmdMatrix3(80).Caption = 80
                cmdMatrix3(83).Caption = 83
                cmdMatrix3(0).BackColor = &HC0FFFF
                cmdMatrix3(3).BackColor = &HC0FFFF
                cmdMatrix3(80).BackColor = &HC0FFFF
                cmdMatrix3(83).BackColor = &HC0FFFF
        Case 4
                cmdMatrix4(0).Caption = 0
                cmdMatrix4(3).Caption = 3
                cmdMatrix4(80).Caption = 80
                cmdMatrix4(83).Caption = 83
                cmdMatrix4(0).BackColor = &HC0FFFF
                cmdMatrix4(3).BackColor = &HC0FFFF
                cmdMatrix4(80).BackColor = &HC0FFFF
                cmdMatrix4(83).BackColor = &HC0FFFF
        End Select
End Select


End Sub


Private Sub cmdD_Click()
txtYOffset.Text = txtYOffset.Text - 0.001
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFire_Click()

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
    
    SEG_Y_DIST = 0
    LICA_XOFF = 0
    LICA_YOFF = 0
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

       
    Dim X As Integer, Y As Integer, k As Integer
    
    Dim SEGMENT_COUNT_1 As Integer
    Dim SEGMENT_COUNT_2 As Integer
    Dim SEGMENT_COUNT_3 As Integer
    Dim SEGMENT_COUNT_4 As Integer
    Dim SEGMENT_COUNT_5 As Integer
    
    For X = 0 To 3
           For Y = 0 To 8
                k = X + (10 * Y)
                    If (cmdMatrix(k).Caption <> "") Then
                            FIRE_MATRIX(1, k) = 1
                            SEGMENT_COUNT_1 = SEGMENT_COUNT_1 + 1
                    Else
                            FIRE_MATRIX(1, k) = 0
                    End If
                    If (cmdMatrix1(k).Caption <> "") Then
                            FIRE_MATRIX(2, k) = 1
                            SEGMENT_COUNT_2 = SEGMENT_COUNT_2 + 1
                    Else
                            FIRE_MATRIX(2, k) = 0
                    End If
                    If (cmdMatrix2(k).Caption <> "") Then
                            FIRE_MATRIX(3, k) = 1
                            SEGMENT_COUNT_3 = SEGMENT_COUNT_3 + 1
                    Else
                            FIRE_MATRIX(3, k) = 0
                    End If
                    If (cmdMatrix3(k).Caption <> "") Then
                            FIRE_MATRIX(4, k) = 1
                            SEGMENT_COUNT_4 = SEGMENT_COUNT_4 + 1
                    Else
                            FIRE_MATRIX(4, k) = 0
                    End If
                    If (cmdMatrix4(k).Caption <> "") Then
                            FIRE_MATRIX(5, k) = 1
                            SEGMENT_COUNT_5 = SEGMENT_COUNT_5 + 1
                    Else
                            FIRE_MATRIX(5, k) = 0
                    End If
            Next Y
    Next X
                                          
    If (Option1.value = True) Then
        SEGMENT_ID = 1
        If (SEGMENT_COUNT_1 = 0) Then
                MsgBox "No Chips Selected", vbInformation, "ITP Tray Laser System"
                Exit Sub
        End If
    End If
    If (Option2.value = True) Then
        SEGMENT_ID = 2
        If (SEGMENT_COUNT_2 = 0) Then
                MsgBox "No Chips Selected", vbInformation, "ITP Tray Laser System"
                Exit Sub
        End If
    End If
    If (Option3.value = True) Then
        SEGMENT_ID = 3
        If (SEGMENT_COUNT_3 = 0) Then
                MsgBox "No Chips Selected", vbInformation, "ITP Tray Laser System"
                Exit Sub
        End If
    End If
    If (Option4.value = True) Then
        If (SEGMENT_COUNT_4 = 0) Then
                MsgBox "No Chips Selected", vbInformation, "ITP Tray Laser System"
                Exit Sub
        End If
        SEGMENT_ID = 4
    End If
    If (Option5.value = True) Then
        If (SEGMENT_COUNT_5 = 0) Then
                MsgBox "No Chips Selected", vbInformation, "ITP Tray Laser System"
                Exit Sub
        End If
        SEGMENT_ID = 5
    End If
    If (OptionAll.value = True) Then
         If (SEGMENT_COUNT_1 + SEGMENT_COUNT_2 + SEGMENT_COUNT_2 + SEGMENT_COUNT_4 + SEGMENT_COUNT_5 = 0) Then
                MsgBox "No Chips Selected", vbInformation, "ITP Tray Laser System"
                Exit Sub
        End If
     End If
    '
    '   All Segment or Just Selected Segment
    '
    cmdFire.Enabled = False
    cmdFire.BackColor = vbButtonFace
    
    If (OptionAll.value = False) Then
            For X = 0 To 3
                   For Y = 0 To 8
                        k = X + (10 * Y)
                        If (FIRE_MATRIX(SEGMENT_ID, k) = 1) Then
                                Fire_Objects (k)
                                MARK_COUNT = MARK_COUNT + 1
                        End If
                    Next Y
            Next X
    End If
        
        
    '*********** PRODUCTION MODE
    
    If (OptionAll.value = True) Then
    
        For SEGMENT_ID = 1 To 5
            '
            'MOVE X AXIS TO NEXT SEGMENT
            '
            MoveToTarget 1, CLng(txtTargetPos(SEGMENT_ID).Text)
            
            lblSegment.Caption = SEGMENT_ID
            lblSegment.Refresh
            '
            'FIRE
            '
            For X = 0 To 3
                   For Y = 0 To 8
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

Private Sub cmdMatrix1_Click(Index As Integer)

If (cmdMatrix1(Index).Caption = "") Then
    cmdMatrix1(Index).Caption = Index
    cmdMatrix1(Index).BackColor = &HC0FFFF
Else
    cmdMatrix1(Index).Caption = ""
    cmdMatrix1(Index).BackColor = &H808080
End If

End Sub

Private Sub cmdMatrix2_Click(Index As Integer)

If (cmdMatrix2(Index).Caption = "") Then
    cmdMatrix2(Index).Caption = Index
    cmdMatrix2(Index).BackColor = &HC0FFFF
Else
    cmdMatrix2(Index).Caption = ""
    cmdMatrix2(Index).BackColor = &H808080
End If

End Sub

Private Sub cmdMatrix3_Click(Index As Integer)

If (cmdMatrix3(Index).Caption = "") Then
    cmdMatrix3(Index).Caption = Index
    cmdMatrix3(Index).BackColor = &HC0FFFF
Else
    cmdMatrix3(Index).Caption = ""
    cmdMatrix3(Index).BackColor = &H808080
End If

End Sub

Private Sub cmdMatrix4_Click(Index As Integer)
If (cmdMatrix4(Index).Caption = "") Then
    cmdMatrix4(Index).Caption = Index
    cmdMatrix4(Index).BackColor = &HC0FFFF
Else
    cmdMatrix4(Index).Caption = ""
    cmdMatrix4(Index).BackColor = &H808080
End If
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
For X = 0 To 3
       For Y = 0 To 8
            k = X + (10 * Y)
            Select Case Index
            Case 0
                    If (cmdMatrix(k).Enabled = True) Then
                        cmdMatrix(k).Caption = ""
                        cmdMatrix(k).BackColor = &H808080
                    End If
            Case 1
                    If (cmdMatrix1(k).Enabled = True) Then
                        cmdMatrix1(k).Caption = ""
                        cmdMatrix1(k).BackColor = &H808080
                    End If
            Case 2
                    If (cmdMatrix2(k).Enabled = True) Then
                        cmdMatrix2(k).Caption = ""
                        cmdMatrix2(k).BackColor = &H808080
                    End If
            Case 3
                    If (cmdMatrix3(k).Enabled = True) Then
                        cmdMatrix3(k).Caption = ""
                        cmdMatrix3(k).BackColor = &H808080
                    End If
            Case 4
                    If (cmdMatrix4(k).Enabled = True) Then
                        cmdMatrix4(k).Caption = ""
                        cmdMatrix4(k).BackColor = &H808080
                    End If
            End Select
            
        Next Y
Next X

End Sub

Private Sub cmdReverse_Click(Index As Integer)

Dim X As Integer, Y As Integer, k As Integer
For X = 0 To 3
       For Y = 0 To 8
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
            Case 1
                    If (cmdMatrix1(k).Enabled = True) Then
                        If (cmdMatrix1(k).Caption = "") Then
                            cmdMatrix1(k).Caption = k
                            cmdMatrix1(k).BackColor = &HC0FFFF
                        Else
                            cmdMatrix1(k).Caption = ""
                            cmdMatrix1(k).BackColor = &H808080
                        End If
                    End If
            Case 2
                    If (cmdMatrix2(k).Enabled = True) Then
                        If (cmdMatrix2(k).Caption = "") Then
                            cmdMatrix2(k).Caption = k
                            cmdMatrix2(k).BackColor = &HC0FFFF
                        Else
                            cmdMatrix2(k).Caption = ""
                            cmdMatrix2(k).BackColor = &H808080
                        End If
                    End If
            Case 3
                    If (cmdMatrix3(k).Enabled = True) Then
                        If (cmdMatrix3(k).Caption = "") Then
                            cmdMatrix3(k).Caption = k
                            cmdMatrix3(k).BackColor = &HC0FFFF
                        Else
                            cmdMatrix3(k).Caption = ""
                            cmdMatrix3(k).BackColor = &H808080
                        End If
                    End If
            End Select

        Next Y
Next X

End Sub

Private Sub cmdRow_Click(Index As Integer)

Dim Y As Integer, k As Integer
Dim i As Integer
 
For Y = 0 To 8
     k = (10 * Y)
     Select Case Index
     Case 0
                If (cmdMatrix(k).Caption <> "") Then
                    For i = 0 To 3
                         cmdMatrix(i + k).Caption = i + k
                         cmdMatrix(i + k).BackColor = &HC0FFFF
                    Next i
                End If
     Case 1
                If (cmdMatrix2(k).Caption <> "") Then
                    For i = 0 To 3
                         cmdMatrix1(i + k).Caption = i + k
                         cmdMatrix1(i + k).BackColor = &HC0FFFF
                    Next i
                End If
     Case 2
                If (cmdMatrix2(k).Caption <> "") Then
                    For i = 0 To 3
                         cmdMatrix2(i + k).Caption = i + k
                         cmdMatrix2(i + k).BackColor = &HC0FFFF
                    Next i
                End If
     Case 3
                If (cmdMatrix3(k).Caption <> "") Then
                    For i = 0 To 3
                         cmdMatrix3(i + k).Caption = i + k
                         cmdMatrix3(i + k).BackColor = &HC0FFFF
                    Next i
                End If
     Case 4
                If (cmdMatrix4(k).Caption <> "") Then
                    For i = 0 To 3
                         cmdMatrix4(i + k).Caption = i + k
                         cmdMatrix4(i + k).BackColor = &HC0FFFF
                    Next i
                End If
     End Select

Next Y
 
End Sub

Private Sub cmdTrayConfig_Click()

Dim X As Integer, Y As Integer, k As Integer

For X = 0 To 3
       For Y = 0 To 8
            k = X + (10 * Y)
            cmdMatrix(k).Enabled = True
            cmdMatrix1(k).Enabled = True
            cmdMatrix2(k).Enabled = True
            cmdMatrix3(k).Enabled = True
            cmdMatrix4(k).Enabled = True
            cmdMatrix(k).BackColor = &H808080
            cmdMatrix1(k).BackColor = &H808080
            cmdMatrix2(k).BackColor = &H808080
            cmdMatrix3(k).BackColor = &H808080
            cmdMatrix4(k).BackColor = &H808080
        Next Y
Next X

If (OptionTray2.value = True) Then

End If
If (OptionTray3.value = True) Then

            For X = 0 To 3
                   For Y = 0 To 8
                        k = X + (10 * Y)
                        Select Case k
                        Case 0 To 3
                        Case 30 To 33
                        Case 50 To 53
                        Case 80 To 83
                        Case Else
                                cmdMatrix(k).Enabled = False
                                cmdMatrix1(k).Enabled = False
                                cmdMatrix2(k).Enabled = False
                                cmdMatrix3(k).Enabled = False
                                cmdMatrix4(k).Enabled = False
                                cmdMatrix(k).BackColor = &H8000000F
                                cmdMatrix1(k).BackColor = &H8000000F
                                cmdMatrix2(k).BackColor = &H8000000F
                                cmdMatrix3(k).BackColor = &H8000000F
                                cmdMatrix4(k).BackColor = &H8000000F
                        End Select
                    Next Y
            Next X
End If
If (OptionTray4.value = True) Then
            For X = 0 To 3
                   For Y = 0 To 8
                        k = X + (10 * Y)
                        Select Case k
                        Case 10 To 13
                        Case 40 To 43
                        Case 70 To 73
                        Case Else
                                cmdMatrix(k).Enabled = False
                                cmdMatrix1(k).Enabled = False
                                cmdMatrix2(k).Enabled = False
                                cmdMatrix3(k).Enabled = False
                                cmdMatrix4(k).Enabled = False
                                cmdMatrix(k).BackColor = &H8000000F
                                cmdMatrix1(k).BackColor = &H8000000F
                                cmdMatrix2(k).BackColor = &H8000000F
                                cmdMatrix3(k).BackColor = &H8000000F
                                cmdMatrix4(k).BackColor = &H8000000F
                        End Select
                    Next Y
            Next X
End If

End Sub

Private Sub cmdU_Click()
txtYOffset.Text = txtYOffset.Text + 0.001
End Sub

Private Sub cmdUpdate_Click()
Data2.UpdateRecord
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

Private Sub CommandUpdateGlobal_Click()

Data2.UpdateRecord

GlobalUpdate

End Sub

Private Sub CommandUpdateHeight_Click()
GlobalHeightUpdate
End Sub

Private Sub Form_Load()

Caption = "ATC " & Get_Title & "   " & ATC_DWG & "         " & ATC_VERSION

Init_Array (MATRIX_ID)

LabelSIZE_LOC_ID.Caption = SIZE_LOC_ID
LabelTRAY_ID.Caption = TRAY_ID
LabelPOWER_ID.Caption = POWER_ID

cmdFire.BackColor = vbGreen

Dim X As Integer, Y As Integer, k As Integer
For X = 0 To 3
       For Y = 0 To 8
            k = X + (10 * Y)
            cmdMatrix(k).Caption = ""
            cmdMatrix1(k).Caption = ""
            cmdMatrix2(k).Caption = ""
            cmdMatrix3(k).Caption = ""
            cmdMatrix4(k).Caption = ""
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

If (TRAY_ID = 2) Then
    OptionTray2.value = True
End If
If (TRAY_ID = 3) Then
    OptionTray3.value = True
End If
If (TRAY_ID = 4) Then
    OptionTray4.value = True
End If
If (TRAY_ID = 12) Then
    OptionTray4.value = True
End If

Data1.DatabaseName = ATC_LASER_BD
Data2.DatabaseName = ATC_LASER_BD
Data3.DatabaseName = ATC_LASER_BD

Dim sSQL As String

sSQL = "SELECT * FROM  [TBL Tray Config] WHERE [TRAY_ID] = " & TRAY_ID
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
' TRAY 301 - 414
'
 
For X = 0 To 3
       For Y = 0 To 8
            k = X + (10 * Y)
            Select Case Y
            Case 0
                    gdLocation(YLOC, k) = 0
            Case 1
                    gdLocation(YLOC, k) = 0.499
            Case 2
                    gdLocation(YLOC, k) = 0.499 + 0.499
            Case 3
                    gdLocation(YLOC, k) = 0.499 + 0.499 + 0.592
            Case 4
                    gdLocation(YLOC, k) = 0.499 + 0.499 + 0.592 + 0.499
            Case 5
                    gdLocation(YLOC, k) = 0.499 + 0.499 + 0.592 + 0.499 + 0.499
            Case 6
                    gdLocation(YLOC, k) = 0.499 + 0.499 + 0.592 + 0.499 + 0.499 + 0.592
            Case 7
                    gdLocation(YLOC, k) = 0.499 + 0.499 + 0.592 + 0.499 + 0.499 + 0.592 + 0.499
            Case 8
                    gdLocation(YLOC, k) = 0.499 + 0.499 + 0.592 + 0.499 + 0.499 + 0.592 + 0.499 + 0.499
            End Select
            gdLocation(XLOC, k) = X * 0.497
        Next Y
Next X


For X = 0 To 3
       For Y = 0 To 8
            k = X + (10 * Y)
            Select Case Y
            Case 0
                    gdLocation(YLOC, k) = 0
            Case 1
                    gdLocation(YLOC, k) = 0.505
            Case 2
                    gdLocation(YLOC, k) = 0.505 + 0.505
            Case 3
                    gdLocation(YLOC, k) = 0.505 + 0.505 + 0.592
            Case 4
                    gdLocation(YLOC, k) = 0.505 + 0.505 + 0.592 + 0.505
            Case 5
                    gdLocation(YLOC, k) = 0.505 + 0.505 + 0.592 + 0.505 + 0.505
            Case 6
                    gdLocation(YLOC, k) = 0.505 + 0.505 + 0.592 + 0.505 + 0.505 + 0.592
            Case 7
                    gdLocation(YLOC, k) = 0.505 + 0.505 + 0.592 + 0.505 + 0.505 + 0.592 + 0.505
            Case 8
                    gdLocation(YLOC, k) = 0.505 + 0.505 + 0.592 + 0.505 + 0.505 + 0.592 + 0.505 + 0.505
            End Select
            gdLocation(XLOC, k) = X * 0.497
        Next Y
Next X



dOffSet(XLOC) = 3 * 0.5 / 2
dOffSet(YLOC) = (gdLocation(YLOC, 80) - gdLocation(YLOC, 0)) / 2
           
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

Private Sub OptionTray2_Click()

cmdTrayConfig_Click

End Sub

Private Sub OptionTray3_Click()

cmdTrayConfig_Click

End Sub

Private Sub OptionTray4_Click()

cmdTrayConfig_Click

End Sub

Private Sub txtTargetPos_GotFocus(Index As Integer)

txtTargetPos(Index).SelStart = 0
txtTargetPos(Index).SelLength = Len(txtTargetPos(Index).Text)

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

End Sub
