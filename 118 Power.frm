VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPower 
   Caption         =   "104 Power & Scale Factors"
   ClientHeight    =   11100
   ClientLeft      =   4530
   ClientTop       =   675
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "118 Power.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11100
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdateRecord2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Update Record Power"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   5160
      Width           =   2775
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
      Left            =   240
      TabIndex        =   88
      Text            =   "ATC PART"
      ToolTipText     =   "ATC PART"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      DataField       =   "CASE "
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
      Index           =   20
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   87
      Text            =   "C"
      ToolTipText     =   "CASE"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Active"
      DataField       =   "ACTIVE"
      DataSource      =   "Data2"
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
      Left            =   2880
      TabIndex        =   86
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Active"
      DataField       =   "ACTIVE"
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
      Left            =   8760
      TabIndex        =   85
      Top             =   3000
      Width           =   1335
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
      Left            =   3360
      TabIndex        =   84
      Text            =   "COATING"
      ToolTipText     =   "COATING"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtXOffset 
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
      Height          =   375
      Left            =   13920
      TabIndex        =   81
      Text            =   "0"
      Top             =   2400
      Width           =   765
   End
   Begin VB.TextBox txtYOffset 
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
      Height          =   375
      Left            =   13920
      TabIndex        =   80
      Text            =   "0"
      Top             =   2880
      Width           =   765
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
      Left            =   12360
      TabIndex        =   71
      Top             =   0
      Width           =   2655
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
         TabIndex        =   79
         Top             =   360
         Width           =   1245
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
         TabIndex        =   78
         Top             =   1620
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
         TabIndex        =   77
         Top             =   1200
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
         TabIndex        =   76
         Top             =   780
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
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   1620
         Visible         =   0   'False
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
         TabIndex        =   73
         Top             =   1200
         Visible         =   0   'False
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
         Left            =   1080
         TabIndex        =   72
         Top             =   780
         Width           =   1245
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
      Height          =   2175
      Index           =   2
      Left            =   10320
      TabIndex        =   64
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton optLogo 
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
         Index           =   5
         Left            =   240
         TabIndex        =   70
         Top             =   2580
         Width           =   2100
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
         TabIndex        =   69
         Top             =   2160
         Width           =   2100
      End
      Begin VB.OptionButton optLogo 
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
         Index           =   3
         Left            =   240
         TabIndex        =   68
         Top             =   1740
         Width           =   900
      End
      Begin VB.OptionButton optLogo 
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
         Index           =   0
         Left            =   240
         TabIndex        =   67
         Top             =   480
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optLogo 
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
         Index           =   1
         Left            =   240
         TabIndex        =   66
         Top             =   1320
         Width           =   900
      End
      Begin VB.OptionButton optLogo 
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
         Index           =   2
         Left            =   240
         TabIndex        =   65
         Top             =   900
         Width           =   900
      End
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
      Left            =   10440
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Timer tmrLaser 
      Interval        =   1000
      Left            =   10440
      Top             =   840
   End
   Begin VB.TextBox txtScale 
      Alignment       =   2  'Center
      DataField       =   "CASE SIZE"
      DataSource      =   "Data1"
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
      Index           =   18
      Left            =   8400
      TabIndex        =   62
      Text            =   "Case"
      ToolTipText     =   "CASE SIZE"
      Top             =   2520
      Width           =   1815
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
      TabIndex        =   61
      Text            =   "Series"
      ToolTipText     =   "SERIES"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Update Record Font"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   10200
      Width           =   2415
   End
   Begin VB.CommandButton cmdRefresh2 
      Caption         =   "Refresh2"
      Height          =   300
      Left            =   4800
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 [TBL Size Location]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Size Location"
      Top             =   3240
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Power"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 [TBL Power]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Power"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   4560
      TabIndex        =   55
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LY4"
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
      Index           =   16
      Left            =   14040
      TabIndex        =   50
      Text            =   "LY4"
      ToolTipText     =   "LY4"
      Top             =   9480
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LY3"
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
      Index           =   14
      Left            =   14040
      TabIndex        =   49
      Text            =   "LY3"
      ToolTipText     =   "LY3"
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LY2"
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
      Index           =   9
      Left            =   14040
      TabIndex        =   48
      Text            =   "LY2"
      ToolTipText     =   "LY2"
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LY1"
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
      Index           =   6
      Left            =   14040
      TabIndex        =   47
      Text            =   "LY1"
      ToolTipText     =   "LY1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LY0"
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
      Index           =   5
      Left            =   14040
      TabIndex        =   46
      Text            =   "LY0"
      ToolTipText     =   "LY0"
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LX4"
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
      Index           =   13
      Left            =   13080
      TabIndex        =   45
      Text            =   "LX4"
      ToolTipText     =   "LX4"
      Top             =   9480
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LX3"
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
      Index           =   12
      Left            =   13080
      TabIndex        =   44
      Text            =   "LX3"
      ToolTipText     =   "LX3"
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LX2"
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
      Index           =   11
      Left            =   13080
      TabIndex        =   43
      Text            =   "LX2"
      ToolTipText     =   "LX2"
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LX1"
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
      Index           =   1
      Left            =   13080
      TabIndex        =   42
      Text            =   "LX1"
      ToolTipText     =   "LX1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      DataField       =   "LX0"
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
      Index           =   15
      Left            =   13080
      TabIndex        =   40
      Text            =   "LX0"
      ToolTipText     =   "LX0"
      Top             =   7440
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [TBL Size Location]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Size Location"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Frame Frame1 
      Caption         =   " Power "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   27
      Top             =   5640
      Width           =   7455
      Begin VB.Frame Frame5 
         Caption         =   " Angle [0 - 359.9] "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3840
         TabIndex        =   108
         Top             =   4080
         Width           =   3375
         Begin VB.TextBox txtAngle 
            Alignment       =   2  'Center
            DataField       =   "ANGLE"
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
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   114
            Text            =   "ANGLE"
            ToolTipText     =   "ANGLE"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdAMM 
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
            TabIndex        =   112
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdAPP 
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
            TabIndex        =   111
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdAP 
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
            TabIndex        =   110
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton cmdAM 
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
            Left            =   720
            TabIndex        =   109
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Z Height "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         TabIndex        =   102
         Top             =   4080
         Width           =   3255
         Begin VB.TextBox txtZHT 
            Alignment       =   2  'Center
            DataField       =   "Z HEIGHT"
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
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   107
            Text            =   "Z"
            ToolTipText     =   "Z HEIGHT"
            Top             =   480
            Width           =   855
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
            TabIndex        =   106
            Top             =   480
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
            TabIndex        =   105
            Top             =   480
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
            TabIndex        =   104
            Top             =   480
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
            TabIndex        =   103
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " QS Mode [0.1 - 50.0] "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   360
         TabIndex        =   32
         Top             =   2760
         Width           =   4335
         Begin VB.CommandButton Command12 
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
            TabIndex        =   100
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton Command11 
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
            TabIndex        =   99
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton Command10 
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
            TabIndex        =   98
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Command9 
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
            TabIndex        =   97
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "KHZ"
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
            Left            =   3480
            TabIndex        =   36
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblQS 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RATE"
            DataField       =   "RATE"
            DataSource      =   "Data2"
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
            Left            =   2520
            TabIndex        =   33
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Marking Speed [0.039 - 24.000] "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   360
         TabIndex        =   30
         Top             =   1560
         Width           =   4335
         Begin VB.CommandButton Command8 
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
            TabIndex        =   96
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command7 
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
            TabIndex        =   95
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command6 
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
            TabIndex        =   94
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command5 
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
            TabIndex        =   93
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "in/sec"
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
            Left            =   3480
            TabIndex        =   35
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblMS 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SPEED"
            DataField       =   "SPEED"
            DataSource      =   "Data2"
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
            Left            =   2520
            TabIndex        =   31
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Current [5.00 - 20.00]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   4335
         Begin VB.CommandButton Command4 
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
            TabIndex        =   92
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command3 
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
            TabIndex        =   91
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command2 
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
            TabIndex        =   90
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command1 
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
            TabIndex        =   89
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Amp"
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
            Left            =   3600
            TabIndex        =   34
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CURRENT"
            DataField       =   "CURRENT"
            DataSource      =   "Data2"
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
            Left            =   2520
            TabIndex        =   29
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.Frame fraLogo 
      Caption         =   " Top Logo Parameters "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   1
      Left            =   7680
      TabIndex        =   17
      Top             =   3480
      Width           =   3255
      Begin VB.TextBox txtScale 
         DataField       =   "LOGO CX TOP"
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
         Index           =   10
         Left            =   2040
         TabIndex        =   21
         Text            =   "CX"
         ToolTipText     =   "LOGO CX TOP"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtSXTop 
         DataField       =   "LOGO SX TOP"
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
         Left            =   2040
         TabIndex        =   20
         Text            =   "SX"
         ToolTipText     =   "LOGO SX TOP"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtScale 
         DataField       =   "LOGO LX TOP"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   19
         Text            =   "LX"
         ToolTipText     =   "LOGO LX TOP"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtScale 
         DataField       =   "LOGO LY TOP"
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
         Index           =   7
         Left            =   2040
         TabIndex        =   18
         Text            =   "LY"
         ToolTipText     =   "LOGO LY TOP"
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Compression X"
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
         Index           =   14
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Scale Height"
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
         Index           =   13
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Location X"
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
         Index           =   12
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Location Y"
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
         Index           =   11
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Logo Height 6"
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
         Index           =   7
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Font Height * Scale Height"
         Top             =   1680
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Font "
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
      Left            =   11160
      TabIndex        =   10
      Top             =   3480
      Width           =   3735
      Begin VB.VScrollBar vsbScale 
         Height          =   1455
         Left            =   2400
         Max             =   800
         Min             =   1
         TabIndex        =   15
         Top             =   1080
         Value           =   1
         Width           =   375
      End
      Begin VB.TextBox txtScale 
         DataField       =   "FONT SCALE"
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
         Index           =   4
         Left            =   2040
         TabIndex        =   11
         Text            =   "SCALE"
         ToolTipText     =   "FONT SCALE"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Font Height"
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
         Index           =   17
         Left            =   600
         TabIndex        =   37
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblFH 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HEIGHT"
         DataField       =   "FONT HEIGHT"
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
         Height          =   375
         Left            =   840
         TabIndex        =   16
         ToolTipText     =   "FONT HEIGHT"
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "100 Full Scale"
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
         Index           =   5
         Left            =   600
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Compression X"
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
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame fraLogo 
      Caption         =   " Sides Logo Parameters "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   7680
      TabIndex        =   1
      Top             =   6840
      Width           =   3255
      Begin VB.TextBox txtScale 
         DataField       =   "LOGO LY SIDE"
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
         Index           =   3
         Left            =   2040
         TabIndex        =   5
         Text            =   "LY"
         ToolTipText     =   "LOGO LY SIDE"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtScale 
         DataField       =   "LOGO LX SIDE"
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
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Tag             =   "LOGO LX SID"
         Text            =   "LX"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtSXSide 
         DataField       =   "LOGO SX SIDE"
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
         Left            =   2040
         TabIndex        =   3
         Text            =   "SX"
         ToolTipText     =   "LOGO SX SIDE"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtScale 
         DataField       =   "LOGO CX SIDE"
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
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Text            =   "CX"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Logo Height 6"
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
         Index           =   6
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Font Height * Scale Height"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Location Y"
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
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Location X"
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
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Scale Height"
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
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Compression X"
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
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdScale 
      Caption         =   "Exit to Main"
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
      Left            =   8400
      TabIndex        =   0
      Top             =   10320
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "118 Power.frx":0CCA
      Height          =   855
      Left            =   0
      TabIndex        =   56
      ToolTipText     =   "FROM [TBL Power] WHERE  POWER_ID"
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1508
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "118 Power.frx":0CDE
      Height          =   1815
      Left            =   7200
      TabIndex        =   58
      ToolTipText     =   "FROM  [TBL Size Location] WHERE SIZE_LOC_ID"
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3201
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      DataField       =   "TBL_ID"
      DataSource      =   "Data2"
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
      Index           =   18
      Left            =   5520
      TabIndex        =   113
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Laser X Offset"
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
      Left            =   12360
      TabIndex        =   83
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Laser Y Offset"
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
      Left            =   12360
      TabIndex        =   82
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Case Size "
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
      Index           =   8
      Left            =   8520
      TabIndex        =   59
      Top             =   2040
      Width           =   1455
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
      Left            =   11520
      TabIndex        =   54
      Top             =   9480
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
      Left            =   11520
      TabIndex        =   53
      Top             =   9000
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
      Left            =   11520
      TabIndex        =   52
      Top             =   8520
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
      Left            =   11520
      TabIndex        =   51
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ATC"
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
      Index           =   21
      Left            =   11520
      TabIndex        =   41
      ToolTipText     =   "ATC Non Mag"
      Top             =   7440
      Width           =   1095
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
      Left            =   14040
      TabIndex        =   39
      Top             =   6960
      Width           =   855
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
      Left            =   13080
      TabIndex        =   38
      Top             =   6960
      Width           =   855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Picture         =   "118 Power.frx":0CF2
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmPower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAM_Click()
txtAngle.Text = txtAngle.Text - 1
End Sub

Private Sub cmdAMM_Click()
txtAngle.Text = txtAngle.Text - 10
End Sub

Private Sub cmdAP_Click()
txtAngle.Text = txtAngle.Text + 1

End Sub

Private Sub cmdAPP_Click()
txtAngle.Text = txtAngle.Text + 10
End Sub

Private Sub cmdFire_Click()

    TRAY_X_OFFSET = Val(txtXOffset.Text)
    TRAY_Y_OFFSET = Val(txtYOffset.Text)
                              
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
    
    If optLogo(0).value = True Then
        LOGO_MODE = 0
    End If
    If optLogo(1).value = True Then
        LOGO_MODE = 1
    End If
    If optLogo(2).value = True Then
        LOGO_MODE = 2
    End If
    If optLogo(3).value = True Then
        LOGO_MODE = 3
    End If
    If optLogo(4).value = True Then
        LOGO_MODE = 4
    End If
    If optLogo(4).value = True Then
        LOGO_MODE = 5
    End If
       
    LASER_TXT1 = UCase(Text1.Text)
    LASER_TXT2 = UCase(Text2.Text)
    LASER_TXT3 = UCase(Text3.Text)
    LASER_TXT4 = UCase(Text4.Text)
    
    SEGMENT_ID = 1
        
    Dim X As Integer, Y As Integer, k As Integer
    For X = 0 To 9
           For Y = 0 To 9
                k = X + (10 * Y)
                FIRE_MATRIX(1, k) = 0
                gdLocation(XLOC, k) = 0
                gdLocation(YLOC, k) = 0
            Next Y
    Next X
    FIRE_MATRIX(1, 0) = 1

    If (OutputDataFileNew = True) Then
        
            cmdFire.Enabled = False
            cmdFire.BackColor = vbButtonFace
            
            Dim iFilenum As Integer
            Dim sFilename As String
                
            sFilename = "C:\ATC\FIRE.TXT"
            iFilenum = FreeFile
            Open sFilename For Output As iFilenum
            Print #iFilenum, 1
            Close iFilenum

    Else
            MsgBox "No Chips Selected", vbInformation, "ATC"
    End If
    
    tmrLaser.Enabled = True
    tmrLaser.Interval = 1000
        
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

Private Sub cmdRefresh_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [TBL_ID],[SERIES],[CASE],[VALUE],[COATING],[ATC PART]," & _
              "format([CURRENT],'0.00')," & _
              "format([SPEED],'0.000')," & _
              "format([RATE],'0.00')," & _
                     "[Z HEIGHT] " & _
       "FROM [TBL Power] " & _
       "WHERE [ACTIVE] = Yes AND " & _
             "[CASE]    ='" & CASE_ID & "' AND " & _
             "[TRAY_ID]  =" & TRAY_ID & " AND " & _
             "[TRAY_TYPE]=" & TRAY_TYPE_ID
                                   
sSQLF = "   ||<Series    |^Case|^Range    |<Coating          |<ATC Part        |>Current |>Speed    |>Rate      |<ZHT   "


Data3.RecordSource = sSQL
Data3.Refresh
 
MSFlexGrid3.FormatString = sSQLF

End Sub

Private Sub cmdRefresh2_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [SIZE_LOC_ID],[CASE SIZE]," & _
              "format([FONT HEIGHT],'0.000') " & _
       "FROM [TBL Size Location]"
                                   
sSQLF = "   ||<Case Size          |Font Height "

Data4.RecordSource = sSQL
Data4.Refresh
 
MSFlexGrid4.FormatString = sSQLF

End Sub

'
Private Sub cmdScale_Click(Index As Integer)

Unload Me
  
End Sub

Private Sub cmdUpdate_Click()

Data1.UpdateRecord

cmdRefresh2_Click
MSFlexGrid4_Click

End Sub

Private Sub cmdUpdateRecord2_Click()

Data2.UpdateRecord

cmdRefresh_Click
MSFlexGrid3_Click

End Sub

Private Sub Command1_Click()
lblCurrent.Caption = lblCurrent.Caption - 0.1
End Sub

Private Sub Command10_Click()
lblQS.Caption = lblQS.Caption + 0.1
End Sub

Private Sub Command11_Click()
lblQS.Caption = lblQS.Caption + 1
End Sub

Private Sub Command12_Click()
lblQS.Caption = lblQS.Caption - 1
End Sub

Private Sub Command2_Click()
lblCurrent.Caption = lblCurrent.Caption + 0.1
End Sub

Private Sub Command3_Click()
lblCurrent.Caption = lblCurrent.Caption + 1
End Sub

Private Sub Command4_Click()
lblCurrent.Caption = lblCurrent.Caption - 1
End Sub

Private Sub Command5_Click()
lblMS.Caption = lblMS.Caption - 1
End Sub

Private Sub Command6_Click()
lblMS.Caption = lblMS.Caption + 1
End Sub

Private Sub Command7_Click()
lblMS.Caption = lblMS.Caption + 0.1
End Sub

Private Sub Command8_Click()
lblMS.Caption = lblMS.Caption - 0.1
End Sub

Private Sub Command9_Click()
lblQS.Caption = lblQS.Caption - 0.1
End Sub

Private Sub Form_Activate()

Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String

sSQL = "SELECT * FROM  [TBL Size Location] WHERE [SIZE_LOC_ID] = " & SIZE_LOC_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
       
    Label1(6).Caption = "Logo Height = " & Val(txtSXSide.Text) * Val(lblFH.Caption)
    Label1(7).Caption = "Logo Height = " & Val(txtSXTop.Text) * Val(lblFH.Caption)
    
    lblFH.Caption = Format(FR_Table.Fields("[FONT HEIGHT]"), "0.000")
    vsbScale.value = FR_Table.Fields("[FONT HEIGHT]") * 1000
    
End If

End Sub
'
'
Private Sub Form_Load()

Caption = "Power & Scale Factors           " & ATC_DWG & "         " & ATC_VERSION

Data1.DatabaseName = ATC_LASER_BD
Data2.DatabaseName = ATC_LASER_BD
Data3.DatabaseName = ATC_LASER_BD
Data4.DatabaseName = ATC_LASER_BD

MSFlexGrid3.Height = 3600

cmdRefresh_Click
MSFlexGrid3_Click

cmdRefresh2_Click
MSFlexGrid4_Click

If (Len(Text1.Text & "X") = 1) Then
    Text1.Text = LASER_TXT1
End If
If (Len(Text2.Text & "X") = 1) Then
    Text2.Text = LASER_TXT2
End If
If (Len(Text3.Text & "X") = 1) Then
    Text3.Text = LASER_TXT3
End If
If (Len(Text4.Text & "X") = 1) Then
    Text4.Text = LASER_TXT4
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub MSFlexGrid3_Click()

MSFlexGrid3.Col = 1
POWER_ID = Val(MSFlexGrid3.Text)
 
MSFlexGrid3.Col = 0
MSFlexGrid3.ColSel = MSFlexGrid3.Cols - 1 '10

Dim sSQL As String

sSQL = "SELECT * FROM [TBL Power] WHERE [TBL_ID] = " & POWER_ID
                                   
Data2.RecordSource = sSQL
Data2.Refresh

End Sub

Private Sub MSFlexGrid4_Click()

MSFlexGrid4.Col = 1
SIZE_LOC_ID = Val(MSFlexGrid4.Text)
 
MSFlexGrid4.Col = 0
MSFlexGrid4.ColSel = MSFlexGrid4.Cols - 1 '10

Dim sSQL As String

sSQL = "SELECT * FROM  [TBL Size Location] WHERE [SIZE_LOC_ID] = " & SIZE_LOC_ID
                                   
Data1.RecordSource = sSQL
Data1.Refresh

Set FR_Database = OpenDatabase(ATC_LASER_BD)
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
       
    Label1(6).Caption = "Logo Height = " & Val(txtSXSide.Text) * Val(lblFH.Caption)
    Label1(7).Caption = "Logo Height = " & Val(txtSXTop.Text) * Val(lblFH.Caption)
    
    lblFH.Caption = Format(FR_Table.Fields("[FONT HEIGHT]"), "0.000")
    vsbScale.value = FR_Table.Fields("[FONT HEIGHT]") * 1000

End If

FR_Table.Close
FR_Database.Close

End Sub

Private Sub tmrLaser_Timer()
Dim iFilenum As Integer
Dim sFilename As String

sFilename = "C:\ATC\LASER.TXT"
iFilenum = FreeFile

Dim iLaserStatus As Integer

Open sFilename For Input As iFilenum
Input #iFilenum, iLaserStatus
Close iFilenum

Select Case iLaserStatus
Case 1
       ' stbSpec.Panels.Item(4).Text = "READY"
        cmdFire.BackColor = vbGreen
        cmdFire.Enabled = True
        cmdFire.SetFocus
        tmrLaser.Enabled = False
Case 0
        'stbSpec.Panels.Item(4).Text = "BUSY"
        cmdFire.Enabled = False
End Select
End Sub

Private Sub txtScale_GotFocus(Index As Integer)
txtScale(Index).SelStart = 0
txtScale(Index).SelLength = Len(txtScale(Index))
End Sub

'
Private Sub vsbScale_Change()

lblFH.Caption = Format(vsbScale.value / 1000, "0.000")

End Sub
