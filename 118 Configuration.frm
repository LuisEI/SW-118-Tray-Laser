VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "118 Configuration Tray Laser"
   ClientHeight    =   11475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11475
   ScaleWidth      =   18600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommandSet 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Fixture"
      Height          =   300
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   10320
      Width           =   1215
   End
   Begin VB.TextBox TextATCPart 
      BackColor       =   &H00FFFFC0&
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
      Left            =   3480
      MaxLength       =   15
      TabIndex        =   77
      Text            =   "100E102FQX"
      ToolTipText     =   " "
      Top             =   10320
      Width           =   1680
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      DataField       =   "SQL_SERIES"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   69
      Text            =   "1"
      ToolTipText     =   "SQL_SERIES"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H00FFFFC0&
      DataField       =   "CAMERA POS"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   67
      Text            =   "1"
      ToolTipText     =   "CAMERA POS"
      Top             =   9240
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      DataField       =   "HEIGHT"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   65
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   8940
      Width           =   735
   End
   Begin VB.CommandButton CommandExcelIn 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Excel In"
      Height          =   300
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   10080
      Width           =   1000
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [TBL Tray Config]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   1440
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.CommandButton cmdRefresh1 
      Caption         =   "Refresh1"
      Height          =   300
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CommandSQL 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SQL"
      Height          =   300
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   10080
      Width           =   1000
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Record"
      Height          =   300
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton CommandExcelOut 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Excel Out"
      Height          =   300
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   10080
      Width           =   1000
   End
   Begin VB.CommandButton cmdRefresh6 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   720
      TabIndex        =   42
      Top             =   10800
      Width           =   1215
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Tray Config"
      Top             =   4920
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL POWER"
      Top             =   5400
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      DataField       =   "TRAY_ID"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      MaxLength       =   2
      TabIndex        =   41
      Text            =   "12"
      ToolTipText     =   "Spacing INDEX"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      DataField       =   "ATC DWG"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   40
      Text            =   "1"
      ToolTipText     =   "ATC DWG"
      Top             =   4725
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "[2]  [MARK PARA],[TRAY_ID],[ORDER]"
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
      Left            =   9480
      TabIndex        =   39
      Top             =   1800
      Width           =   4035
   End
   Begin VB.OptionButton Option1 
      Caption         =   "[1] ORDER BY [TRAY_ID],[ORDER]"
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
      Left            =   9480
      TabIndex        =   38
      Top             =   1440
      Value           =   -1  'True
      Width           =   3675
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFC0&
      DataField       =   "ATC PART"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   37
      Text            =   "1"
      ToolTipText     =   "ATC PART"
      Top             =   5010
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFC0&
      DataField       =   "COATING"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   36
      Text            =   "1"
      ToolTipText     =   "COATING"
      Top             =   5295
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFC0&
      DataField       =   "SERIES"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   35
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   5580
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFC0&
      DataField       =   "CASE "
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   34
      Text            =   "1"
      ToolTipText     =   "CASE"
      Top             =   6225
      Width           =   375
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFC0&
      DataField       =   "VALUE"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   33
      Text            =   "1"
      ToolTipText     =   "Spacing INDEX"
      Top             =   6510
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FFFFC0&
      DataField       =   "DV MIN"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   32
      Text            =   "1"
      ToolTipText     =   "DV MIN"
      Top             =   6795
      Width           =   1335
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFC0&
      DataField       =   "DV MAX"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   31
      Text            =   "1"
      ToolTipText     =   "DV MAX"
      Top             =   7080
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "[3]  TRAY_ID"
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
      Left            =   9480
      TabIndex        =   30
      Top             =   2160
      Width           =   2000
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H00FFFFC0&
      DataField       =   "STAR"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   15600
      TabIndex        =   29
      Text            =   "1"
      ToolTipText     =   "CASE"
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 9"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   28
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 10 MAG"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   27
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   8085
      Width           =   1935
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 10 NON"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   26
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   8370
      Width           =   1935
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFC0&
      DataField       =   "POS 11"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16200
      TabIndex        =   25
      Text            =   "1"
      ToolTipText     =   "SERIES"
      Top             =   8655
      Width           =   735
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H00FFFFC0&
      DataField       =   "ORDER"
      DataSource      =   "Data7"
      Height          =   285
      Left            =   16920
      TabIndex        =   24
      Text            =   "1"
      ToolTipText     =   "ORDER"
      Top             =   7440
      Width           =   375
   End
   Begin VB.Frame Frame3 
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
      Height          =   1335
      Left            =   5280
      TabIndex        =   16
      Top             =   1320
      Width           =   3975
      Begin VB.TextBox txtBOARD_ID 
         Alignment       =   2  'Center
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   2760
         TabIndex        =   22
         Text            =   "1"
         Top             =   840
         Width           =   480
      End
      Begin VB.TextBox TextINITIALIZE_TRAY 
         Alignment       =   2  'Center
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   2760
         TabIndex        =   19
         Text            =   "1"
         Top             =   360
         Width           =   480
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1"
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   400
      End
      Begin VB.CommandButton Command7 
         Caption         =   "0"
         Height          =   300
         Left            =   2280
         TabIndex        =   17
         Top             =   360
         Width           =   400
      End
      Begin VB.Label lblInfo 
         Caption         =   "National Instruments BOARD_ID"
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   2475
      End
      Begin VB.Label lblInfo 
         Caption         =   "INITIALIZE_TRAY"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.Frame fraDB 
      Caption         =   " Data Base Mode /  OP_MODE "
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
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4575
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   1200
         TabIndex        =   76
         Text            =   "XX"
         Top             =   840
         Width           =   400
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NY"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   840
         Width           =   400
      End
      Begin VB.CommandButton Command2 
         Caption         =   "JR"
         Height          =   300
         Left            =   720
         TabIndex        =   74
         Top             =   840
         Width           =   400
      End
      Begin VB.CommandButton Command4 
         Caption         =   "WS [1]"
         Height          =   300
         Left            =   2640
         TabIndex        =   73
         Top             =   840
         Width           =   800
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Test [0]"
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   840
         Width           =   800
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   3600
         TabIndex        =   71
         Text            =   "1"
         Top             =   840
         Width           =   400
      End
      Begin VB.CommandButton Command6 
         Caption         =   "JR: 4"
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   700
      End
      Begin VB.CommandButton Command1 
         Caption         =   "FILE: 2"
         Height          =   300
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   700
      End
      Begin VB.CommandButton cmdLocal 
         BackColor       =   &H00C0FFC0&
         Caption         =   "LCL : 1"
         Height          =   300
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   700
      End
      Begin VB.CommandButton cmdRemote 
         Caption         =   "NY: 0"
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   700
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   3720
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Save File : 118 Configuration.TXT"
      Height          =   300
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2595
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
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
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "XXXX"
      Top             =   240
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
      Height          =   7215
      Left            =   120
      TabIndex        =   43
      ToolTipText     =   "FROM [TBL Power]"
      Top             =   2760
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   12726
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   14880
      TabIndex        =   63
      ToolTipText     =   "FROM  [TBL Tray Config]"
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7011
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
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
   Begin VB.Label lblInfo 
      Caption         =   "Term Style"
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
      Left            =   9240
      TabIndex        =   85
      Top             =   10320
      Width           =   1035
   End
   Begin VB.Label LabelTS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXX"
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
      Left            =   10320
      TabIndex        =   84
      Top             =   10320
      Width           =   720
   End
   Begin VB.Label LabelSeriesCase 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXX"
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
      Left            =   6600
      TabIndex        =   83
      Top             =   10320
      Width           =   720
   End
   Begin VB.Label LabelDV_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXX"
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
      Left            =   8160
      TabIndex        =   82
      Top             =   10320
      Width           =   720
   End
   Begin VB.Label lblInfo 
      Caption         =   "SERIES_ID"
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
      Left            =   5400
      TabIndex        =   81
      Top             =   10320
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Caption         =   "DV_ID"
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
      Left            =   7440
      TabIndex        =   80
      Top             =   10320
      Width           =   675
   End
   Begin VB.Label lblInfo 
      Caption         =   "ATC Part :"
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
      Left            =   2160
      TabIndex        =   78
      Top             =   10320
      Width           =   1275
   End
   Begin VB.Label Label9 
      Caption         =   "SERIES"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   70
      Top             =   5880
      Width           =   1200
   End
   Begin VB.Label Label60 
      Caption         =   "CAMERA POS"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   68
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Chip HEIGHT"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   66
      Top             =   8940
      Width           =   1200
   End
   Begin VB.Label Label13 
      Caption         =   "TRAY_ID [1-12]"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   58
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "ATC Dwg"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   57
      Top             =   6405
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "ATC PART"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   56
      Top             =   5010
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "COATING"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   55
      Top             =   5295
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "SERIES"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   54
      Top             =   5580
      Width           =   1200
   End
   Begin VB.Label Label24 
      Caption         =   "CASE"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   53
      Top             =   6225
      Width           =   1200
   End
   Begin VB.Label Label25 
      Caption         =   "Value Range"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   52
      Top             =   6510
      Width           =   1200
   End
   Begin VB.Label Label26 
      Caption         =   "DV MIN"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   51
      Top             =   6795
      Width           =   1200
   End
   Begin VB.Label Label27 
      Caption         =   "DV MAX"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   50
      Top             =   7080
      Width           =   1200
   End
   Begin VB.Label Label54 
      Caption         =   "STAR"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   49
      Top             =   7440
      Width           =   600
   End
   Begin VB.Label Label55 
      Caption         =   "POS 9"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   48
      Top             =   7800
      Width           =   1200
   End
   Begin VB.Label Label56 
      Caption         =   "POS 10 MAG"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   47
      Top             =   8085
      Width           =   1200
   End
   Begin VB.Label Label57 
      Caption         =   "POS 10 NON"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   46
      Top             =   8370
      Width           =   1200
   End
   Begin VB.Label Label58 
      Caption         =   "POS 11 (Mark)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   14880
      TabIndex        =   45
      Top             =   8655
      Width           =   1200
   End
   Begin VB.Label Label59 
      Caption         =   "Order"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   16200
      TabIndex        =   44
      Top             =   7440
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\MARKER\JOB\ATC DPSS ROT.WLJ"
      Height          =   300
      Left            =   11520
      TabIndex        =   21
      Top             =   960
      Width           =   3000
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\MARKER\JOB\ATC DPSS.WLJ"
      Height          =   300
      Left            =   11520
      TabIndex        =   15
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\MARKER\JOB\LOGO.DWG"
      Height          =   300
      Left            =   11520
      TabIndex        =   14
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OEE SPM JR.MDB"
      Height          =   300
      Left            =   8400
      TabIndex        =   13
      Top             =   960
      Width           =   3000
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ATC Electrical Test.MDB"
      Height          =   300
      Left            =   8400
      TabIndex        =   12
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Work Folder C:\ATC"
      Height          =   300
      Left            =   5280
      TabIndex        =   10
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WO SCHED MASTER.MDB"
      Height          =   300
      Left            =   8400
      TabIndex        =   9
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "118 Configuration.TXT"
      Height          =   300
      Left            =   5280
      TabIndex        =   8
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "118 LASER MATRIX.MDB"
      Height          =   300
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   3000
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLocal_Click()
Text1.Text = 1
End Sub

Private Sub cmdRefresh1_Click()
Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TRAY_ID],[CASE],[TITLE],[ROWS]& ' X ' & [COLS],[ATC DWG]" & _
        "FROM  [TBL Tray Config]"
 

sSQLF = "   |^T_ID|^Case|<Tray Title                          |^Row X Col |^ATC DWG  "

Data1.RecordSource = sSQL
Data1.Refresh
 
MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdRefresh6_Click()

Dim sSQL As String
Dim sSQLF As String

If (Option1.value = True) Then
        sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER],[ORD NEW]," & _
                     "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                     "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11],[MARK PARA],format([HEIGHT],'0.000'),[Z HEIGHT],[CAMERA POS] " & _
             "FROM [TBL Power] WHERE "
        
        sSQL = sSQL & " [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ORDER]"
        
        sSQLF = "   |^T_ID|^PID |<       |<       |<ATC Part  "
        
        sSQLF = sSQLF & "   |<Coating Type|^Series             |^Case|^Value Range|^DP   "
        
        sSQLF = sSQLF & "|<POS 9                        |<POS 10          |<POS 10                |POS 11 |^     |^ZHT       |>ZHT    |>CP       "
     
End If

If (Option2.value = True) Then

        sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER],[ORD NEW]," & _
                      "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                      "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11],[MARK PARA],format([HEIGHT],'0.000') " & _
              "FROM [TBL Power] WHERE "
        
        sSQL = sSQL & " [ACTIVE] = Yes  ORDER BY [MARK PARA],[TRAY_ID],[ORDER]"
        
        sSQLF = "   |^T_ID|^PR_ID|<       |<       |<ATC Part  "
        
        sSQLF = sSQLF & "   |<Coating Type|^Series                |^Case|^Value Range|^DP   |<POS 9                        |<POS 10          |<POS 10                |POS 11 |^     |^ZHT           "

End If


If (Option3.value = True) Then

    sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER],[ORD NEW]," & _
                  "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                  "[MARK PARA],format([HEIGHT],'0.000')" & _
          "FROM [TBL Power] WHERE "

    sSQL = sSQL & " [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ORDER]"

    sSQLF = "   |^T_ID|^PR_ID|<       |<       |<ATC Part  "
    sSQLF = sSQLF & "   |<Coating Type|^Series                |^Case|^Value Range|^DP    |^     |^ZHT           "

End If

Data6.RecordSource = sSQL
Data6.Refresh
 
MSFlexGrid6.FormatString = sSQLF

End Sub

Private Sub cmdRemote_Click()
Text1.Text = 0
End Sub

Private Sub cmdSave_Click()

Dim iAns As Integer
iAns = MsgBox("Save Configuration", vbYesNo, "ATC Juarez Mexico")
If (iAns = vbYes) Then

        BOARD_ID = Val(txtBOARD_ID.Text)
        OP_MODE = Val(Text6.Text)
        LOCATION_ID = Text2.Text
        DataBase_MODE = Val(Text1.Text)
        INITIALIZE_TRAY = Val(TextINITIALIZE_TRAY.Text)

        Configuration (FWRITE)
Else
   
End If

End Sub

Private Sub Command1_Click()
Text1.Text = 2
End Sub

Private Sub Command4_Click()
Text6.Text = 1
End Sub

Private Sub Command5_Click()
Text6.Text = 0
End Sub

Private Sub Command6_Click()
Text1.Text = 4
End Sub

Private Sub Command7_Click()
TextINITIALIZE_TRAY.Text = 0
End Sub

Private Sub Command8_Click()
TextINITIALIZE_TRAY.Text = 1
End Sub

Private Sub Command9_Click()
Data7.UpdateRecord
cmdRefresh6_Click
End Sub

Private Sub CommandExcelIn_Click()

Screen.MousePointer = vbHourglass
                
Dim wbWorld As Object, shtWorld As Object
Dim tSheet As Object

Dim sFilename As String
sFilename = "C:\Documents and Settings\rsoulagnet\My Documents\TRAY LASER.xls"
 
Set shtWorld = GetObject(sFilename)
shtWorld.Application.Visible = False
    
Set wbWorld = shtWorld.Application.Workbooks("TRAY LASER.xls")
Set tSheet = wbWorld.Sheets("sheet1")

Dim iRow As Integer
Dim COUNT As Long

Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String
             
sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER],[ORD NEW]," & _
              "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
              "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11],[MARK PARA],[HEIGHT] " & _
      "FROM [TBL Power] WHERE "

sSQL = sSQL & " [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ORDER]"

Set FR_Table = FR_Database.OpenRecordset(sSQL)

iRow = 3
Do Until FR_Table.EOF
        FR_Table.Edit
        FR_Table.Fields("[MARK PARA]") = tSheet.Cells(iRow, 17).value
        FR_Table.Fields("[HEIGHT]") = tSheet.Cells(iRow, 18).value
        FR_Table.Update
        FR_Table.MoveNext
       iRow = iRow + 1
Loop
 
FR_Database.Close

shtWorld.Application.Quit
Set shtWorld = Nothing

Screen.MousePointer = vbDefault

End Sub

Private Sub CommandExcelOut_Click()

Dim objExcel As Object
Set objExcel = CreateObject("EXCEL.SHEET")
objExcel.Application.Visible = True

Screen.MousePointer = vbHourglass
 
Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TBL Power].[TRAY_ID]                                        AS [SQL 1]," & _
              "[TBL Tray Config].[CASE]                                     AS [SQL 3]," & _
              "[TBL Tray Config].[TITLE]                                    AS [SQL 4]," & _
              "[TBL Tray Config].[ROWS]& ' X ' & [TBL Tray Config].[COLS]   AS [SQL 5]," & _
              "[TBL Tray Config].[ATC DWG]                                  AS [SQL 2]," & _
              "[TBL Power].[TBL_ID]                                         AS [SQL TBL_ID]," & _
              "[TBL Power].[ORDER]," & _
              "[TBL Power].[ORD NEW]," & _
              "[TBL Power].[ATC PART]                                       AS [SQL ATC PART]," & _
              "[TBL Power].[COATING]                                        AS [SQL COATING]," & _
              "[TBL Power].[SERIES]                                         AS [SQL SERIES]," & _
              "[TBL Power].[CASE]                                           AS [SQL CASE]," & _
              "[TBL Power].[VALUE]                                          AS [SQL VALUE]," & _
              "[TBL Power].[STAR]," & _
              "[TBL Power].[POS 9]," & _
              "[TBL Power].[POS 10 MAG]," & _
              "[TBL Power].[POS 10 NON]," & _
              "[TBL Power].[POS 11]," & _
              "[TBL Power].[MARK PARA]                                      AS [SQL MARK MASTER]," & _
       "format([TBL Power].[HEIGHT],'0.000') " & _
    "FROM [TBL Power],[TBL Tray Config] " & _
    "WHERE [TBL Power].[TRAY_ID] = [TBL Tray Config].[TRAY_ID] AND "

sSQL = sSQL & " [TBL Power].[ACTIVE] = Yes  ORDER BY [TBL Power].[TRAY_ID],[TBL Power].[ORDER]"

sSQLF = "   |^T_ID|^PR_ID|<       |<       |<ATC Part  "

sSQLF = sSQLF & "   |<Coating Type|^Series                |^Case|^Value Range|^DP   |<POS 9                        |<POS 10          |<POS 10                |POS 11 |^     |^ZHT           "


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
objExcel.Application.Cells(1, 6).value = "GPWR_ID"
objExcel.Application.Cells(1, 7).value = "MASTER_ID"

objExcel.Application.Cells(1, 8).value = "POWER_ID"
objExcel.Application.Cells(1, 9).value = "ATC PART"
objExcel.Application.Cells(1, 10).value = "COATING"
objExcel.Application.Cells(1, 11).value = "SERIES"
objExcel.Application.Cells(1, 12).value = "CASE"
objExcel.Application.Cells(1, 13).value = "VALUE"

iRow = iRow + 1
Do Until FR_Table.EOF
        objExcel.Application.Cells(iRow, 1).value = FR_Table.Fields("[SQL 1]")
        objExcel.Application.Cells(iRow, 2).value = FR_Table.Fields("[SQL 2]")
        objExcel.Application.Cells(iRow, 3).value = FR_Table.Fields("[SQL 3]")
        objExcel.Application.Cells(iRow, 4).value = FR_Table.Fields("[SQL 4]")
        objExcel.Application.Cells(iRow, 5).value = FR_Table.Fields("[SQL 5]")
        objExcel.Application.Cells(iRow, 6).value = FR_Table.Fields("[SQL MARK MASTER]")
        objExcel.Application.Cells(iRow, 7).value = ""
         
        objExcel.Application.Cells(iRow, 8).value = FR_Table.Fields("[SQL TBL_ID]")
        objExcel.Application.Cells(iRow, 9).value = FR_Table.Fields("[SQL ATC PART]")
        objExcel.Application.Cells(iRow, 10).value = FR_Table.Fields("[SQL COATING]")
        objExcel.Application.Cells(iRow, 11).value = FR_Table.Fields("[SQL SERIES]")
        objExcel.Application.Cells(iRow, 12).value = FR_Table.Fields("[SQL CASE]")
        objExcel.Application.Cells(iRow, 13).value = FR_Table.Fields("[SQL VALUE]")
        iRow = iRow + 1
        FR_Table.MoveNext
Loop
 
FR_Database.Close
                                                                                                     
Dim sFile As String
sFile = "C:\ATC\" & "NEW PARAMETERS.XLS"
                                                                
objExcel.SaveAs sFile
objExcel.Application.Quit
Set objExcel = Nothing
 
Screen.MousePointer = vbDefault
MsgBox "Tray Laser " & sFile, vbInformation, "Excel Format Download"

End Sub

Private Sub CommandSet_Click()

ValidPartNew (TextATCPart.Text)

LabelSeriesCase.Caption = Mid(ATC_PART_ID, 1, 4)
LabelDV_ID.Caption = DV_ID
LabelTS.Caption = Mid(ATC_PART_ID, 9, 2)

Dim sSQL As String
Dim sSQLF As String

Tray_Power_Lookup ATC_PART_ID

If (Option1.value = True) Then
        sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER],[ORD NEW]," & _
                     "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                     "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11],[MARK PARA],format([HEIGHT],'0.000'),[Z HEIGHT],[CAMERA POS] " & _
             "FROM [TBL Power] WHERE "
        sSQL = sSQL & "[PAGE]= 1 AND [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ORDER]"
        
        sSQLF = "   |^T_ID|^PID |<       |<       |<ATC Part  "
        sSQLF = sSQLF & "   |<Coating Type|^Series             |^Case|^Value Range|^DP   "
        sSQLF = sSQLF & "|<POS 9                        |<POS 10          |<POS 10                |POS 11 |^     |^ZHT       |>ZHT    |>CP       "
End If

If (Option2.value = True) Then
        sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER],[ORD NEW]," & _
                      "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                      "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11],[MARK PARA],format([HEIGHT],'0.000') " & _
              "FROM [TBL Power] WHERE "
        sSQL = sSQL & " [PAGE]= 1 AND "
        sSQL = sSQL & " [ACTIVE] = Yes  ORDER BY [MARK PARA],[TRAY_ID],[ORDER]"
        
        sSQLF = "   |^T_ID|^PR_ID|<       |<       |<ATC Part  "
        sSQLF = sSQLF & "   |<Coating Type|^Series                |^Case|^Value Range|^DP   |<POS 9                        |<POS 10          |<POS 10                |POS 11 |^     |^ZHT           "
End If

If (Option3.value = True) Then
    sSQL = "SELECT [TRAY_ID],[TBL_ID],[ORDER],[ORD NEW]," & _
                  "[ATC PART],[COATING],[SERIES],[CASE],[VALUE],[STAR]," & _
                  "[MARK PARA],format([HEIGHT],'0.000')" & _
          "FROM [TBL Power] WHERE "

    sSQL = sSQL & " [PAGE]= 1 AND "
    sSQL = sSQL & " [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ORDER]"

    sSQLF = "   |^T_ID|^PR_ID|<       |<       |<ATC Part  "
    sSQLF = sSQLF & "   |<Coating Type|^Series                |^Case|^Value Range|^DP    |^     |^ZHT           "
End If

Data6.RecordSource = sSQL
Data6.Refresh
 
MSFlexGrid6.FormatString = sSQLF

End Sub

Private Sub CommandSQL_Click()

Set FR_Database = OpenDatabase(ATC_LASER_BD)
 
Dim sSQL As String
    
sSQL = "SELECT * FROM [TBL Power] WHERE  [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ORDER]"
    
Set FR_Table = FR_Database.OpenRecordset(sSQL)
Dim COUNT As Integer
 
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        COUNT = COUNT + 1
        FR_Table.Edit
        FR_Table.Fields("[ORD NEW]") = COUNT
        FR_Table.Update
        FR_Table.MoveNext
    Loop
End If
FR_Table.Close
FR_Database.Close
    
MsgBox "Complete", vbInformation, "ATC"
    
End Sub

Private Sub Form_Load()

Caption = ATC_DWG & "  Configuration Tray Laser   " & ATC_VERSION

txtBOARD_ID.Text = BOARD_ID
Text6.Text = OP_MODE
Text2.Text = LOCATION_ID
Text1.Text = DataBase_MODE

TextINITIALIZE_TRAY.Text = INITIALIZE_TRAY

MSFlexGrid6.Left = 0

Data1.DatabaseName = ATC_LASER_BD

Data6.DatabaseName = ATC_LASER_BD
Data7.DatabaseName = ATC_LASER_BD
cmdRefresh6_Click

cmdRefresh1_Click

TextATCPart.Text = "710E361GA6XJJ"

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

cmdRefresh6_Click

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
End Sub

Private Sub Option2_Click()
cmdRefresh6_Click
End Sub

Private Sub Option3_Click()
cmdRefresh6_Click
End Sub

Private Sub Option4_Click()
cmdRefresh6_Click
End Sub

Private Sub Option5_Click()
cmdRefresh6_Click
End Sub

Private Sub Text22_GotFocus()
Text22.SelStart = 0
Text22.SelLength = Len(Text22)
End Sub

Private Sub TextATCPart_GotFocus()
TextATCPart.SelStart = 0
TextATCPart.SelLength = Len(TextATCPart)
End Sub

Private Sub TextATCPart_LostFocus()
TextATCPart.Text = UCase(TextATCPart.Text)
End Sub
