VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkSheet2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "104 OEE Work Sheet"
   ClientHeight    =   11850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16665
   ControlBox      =   0   'False
   Icon            =   "118 WorkSheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11850
   ScaleWidth      =   16665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data6 
      Caption         =   "Data6 FROM [FIXTURE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   53
      Text            =   "XXXX"
      Top             =   11160
      Width           =   735
   End
   Begin VB.Frame fraWS 
      Caption         =   "Code_ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   4575
      Begin VB.CommandButton cmdUpdateRecord 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Update Record"
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
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2880
         Width           =   1875
      End
      Begin VB.TextBox txtRestock 
         BackColor       =   &H00C0FFC0&
         DataField       =   "RESTOCK"
         DataSource      =   "Data4"
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
         Left            =   3480
         TabIndex        =   37
         Text            =   "12345"
         ToolTipText     =   "Restock"
         Top             =   2280
         Width           =   825
      End
      Begin VB.CommandButton cmdStopTime 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Stop Time"
         Height          =   300
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3840
         Width           =   1800
      End
      Begin VB.TextBox txtATCPart 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ATC PART"
         DataSource      =   "Data4"
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   35
         ToolTipText     =   "Test for Valid Tolerance on exit field"
         Top             =   840
         Width           =   2280
      End
      Begin VB.TextBox txtTotalTime 
         BackColor       =   &H00FFFFFF&
         DataField       =   "TOTAL TIME"
         DataSource      =   "Data4"
         Height          =   300
         Left            =   2400
         TabIndex        =   34
         Text            =   "0"
         ToolTipText     =   "Total time"
         Top             =   3840
         Width           =   600
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Delete"
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
         TabIndex        =   33
         Top             =   2880
         Width           =   1875
      End
      Begin VB.TextBox txtDefects 
         BackColor       =   &H00C0FFC0&
         DataField       =   "REJECTS"
         DataSource      =   "Data4"
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
         Left            =   3480
         TabIndex        =   32
         Text            =   "12345"
         ToolTipText     =   "Defects"
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox txtOrderQty 
         BackColor       =   &H00FFFFC0&
         DataField       =   "QUANTITY"
         DataSource      =   "Data4"
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
         Left            =   1560
         TabIndex        =   31
         Text            =   "12345"
         ToolTipText     =   "Units Produced"
         Top             =   2280
         Width           =   825
      End
      Begin VB.TextBox txtWorkOrder 
         BackColor       =   &H00FFFFC0&
         DataField       =   "WORK ORDER"
         DataSource      =   "Data4"
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   30
         Text            =   "123456789012"
         Top             =   360
         Width           =   2280
      End
      Begin VB.TextBox txtLot 
         BackColor       =   &H00FFFFC0&
         DataField       =   "LOT NUM"
         DataSource      =   "Data4"
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   29
         Text            =   "1234567890"
         Top             =   1320
         Width           =   2280
      End
      Begin VB.TextBox txtSQ 
         BackColor       =   &H00FFFFC0&
         DataField       =   "START QTY"
         DataSource      =   "Data4"
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
         Left            =   1560
         TabIndex        =   28
         Text            =   "12345"
         ToolTipText     =   "Start Qty [START QTY]"
         Top             =   1800
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "START TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   3360
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "h:mm AM/PM"
         Format          =   52887554
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "STOP TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   2400
         TabIndex        =   48
         Top             =   3360
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "h:mm AM/PM"
         Format          =   52887554
         CurrentDate     =   38117
      End
      Begin VB.Label lblInfo 
         Caption         =   "Restock :"
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
         Left            =   2400
         TabIndex        =   46
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Caption         =   "ATC Part :"
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
         TabIndex        =   45
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "Rejects :"
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
         Left            =   2400
         TabIndex        =   44
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Caption         =   "Quantity:"
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
         TabIndex        =   43
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "W.O./Lot#:"
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
         TabIndex        =   42
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "Total Time (m):"
         Height          =   300
         Index           =   4
         Left            =   3240
         TabIndex        =   41
         Top             =   3840
         Width           =   1065
      End
      Begin VB.Label lblInfo 
         Caption         =   "Lot Number:"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "Start Qty:"
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
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   1035
      End
   End
   Begin VB.Frame fraDateSelect 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   20
      Top             =   10800
      Width           =   5895
      Begin VB.CommandButton cmdRefresh1 
         Caption         =   "Refresh"
         Height          =   300
         Left            =   4575
         TabIndex        =   24
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Day  <<"
         Height          =   300
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Day  >>"
         Height          =   300
         Left            =   2565
         TabIndex        =   22
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   300
         Left            =   3570
         TabIndex        =   21
         Top             =   360
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1300
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "h:mm tt"
         Format          =   52887553
         CurrentDate     =   38117
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data5 QTY FROM [DEFECTS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.CommandButton cmdAddDefect 
      Caption         =   "Add  Defect to List >>"
      Height          =   300
      Left            =   5520
      TabIndex        =   19
      Top             =   8160
      Width           =   2880
   End
   Begin VB.Frame fraQTY 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   12480
      TabIndex        =   15
      Top             =   8160
      Width           =   2775
      Begin VB.CommandButton cmdClear 
         Caption         =   "CLR Qty=0"
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1560
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Refresh"
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
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1560
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         DataField       =   "QTY"
         DataSource      =   "Data5"
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
         Left            =   1800
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "Defects"
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdRefresh3 
      Caption         =   "<< Refresh3"
      Height          =   300
      Left            =   10440
      TabIndex        =   13
      Top             =   8280
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 FROM [WORK SHEET],[DEFECT LIST],[DEFECTS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [Defect List]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7440
      Visible         =   0   'False
      Width           =   4140
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "118 WorkSheet.frx":0CCA
      Height          =   975
      Left            =   4920
      TabIndex        =   11
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      BackColorSel    =   16776960
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit to Main"
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
      TabIndex        =   0
      Top             =   11040
      Width           =   2400
   End
   Begin VB.CommandButton cmdCode 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Planned Downtime"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton cmdCode 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Unplanned Downtime"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1980
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton cmdCode 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Code 0"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [WORK SHEET]"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\OEE SPM.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "WORK SHEET"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.CommandButton cmdRefreshDisplay 
      Caption         =   "Refresh Display"
      Height          =   300
      Left            =   8040
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [WORK SHEET]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   3900
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "118 WorkSheet.frx":0CDE
      Height          =   735
      Left            =   5520
      TabIndex        =   12
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      BackColorSel    =   16711680
      AllowBigSelection=   0   'False
      SelectionMode   =   1
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "118 WorkSheet.frx":0CF2
      Height          =   735
      Left            =   8640
      TabIndex        =   14
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      BackColorSel    =   16711680
      AllowBigSelection=   0   'False
      SelectionMode   =   1
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
      Bindings        =   "118 WorkSheet.frx":0D06
      Height          =   2655
      Left            =   240
      TabIndex        =   56
      ToolTipText     =   "FROM [TBL TRAY CONFIG]"
      Top             =   8400
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4683
      _Version        =   393216
      BackColorSel    =   16776960
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
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
   Begin VB.Label LabeMachine_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID"
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
      Left            =   120
      TabIndex        =   58
      Top             =   600
      Width           =   705
   End
   Begin VB.Label LabelLogo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2040
      TabIndex        =   57
      Top             =   7920
      Width           =   2385
   End
   Begin VB.Label Label2 
      Caption         =   "Mag : Side Logo"
      Height          =   300
      Left            =   240
      TabIndex        =   55
      Top             =   7500
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Non Mag : ATC Logo (top)"
      Height          =   300
      Left            =   240
      TabIndex        =   54
      Top             =   7200
      Width           =   2115
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   12360
      TabIndex        =   52
      Top             =   10320
      Width           =   1995
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default: "
      Height          =   300
      Left            =   14400
      TabIndex        =   51
      Top             =   9960
      Width           =   1995
   End
   Begin VB.Label lblDate2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      Height          =   300
      Left            =   12360
      TabIndex        =   50
      Top             =   9960
      Width           =   1995
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
      Height          =   300
      Left            =   14400
      TabIndex        =   49
      Top             =   10320
      Width           =   1995
   End
   Begin VB.Label lblNote 
      Caption         =   "Note : ATC Part "
      Height          =   600
      Left            =   2400
      TabIndex        =   26
      Top             =   7200
      Width           =   2235
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   120
      Picture         =   "118 WorkSheet.frx":0D1A
      Top             =   11160
      Width           =   4170
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "402"
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
      Left            =   150
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "401"
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
      Left            =   150
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblCode 
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
      Height          =   360
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblMachine 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblMachine"
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
      TabIndex        =   7
      Top             =   600
      Width           =   2985
   End
   Begin VB.Label txtShift 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shift"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label txtOperator 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OPERATOR"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2745
   End
End
Attribute VB_Name = "frmWorkSheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

End Sub

Private Sub cmdAddDefect_Click()

If (DEFECT_ID = 0) Then
    Exit Sub
End If

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT * FROM [DEFECTS] " & _
       "WHERE [DEFECT_ID]=" & DEFECT_ID & " AND [WS_ID]=" & WS_ID
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount = 0) Then
    FR_Table.AddNew
    FR_Table.Fields("[WS_ID]") = WS_ID
    FR_Table.Fields("[DEFECT_ID]") = DEFECT_ID
    FR_Table.Fields("[QTY]") = 0
    FR_Table.Update
End If

FR_Table.Close
FR_Database.Close

cmdRefresh3_Click

End Sub

 
Private Sub cmdClear_Click()

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT * FROM [DEFECTS] WHERE [WS_ID]=" & WS_ID & " AND [QTY] = 0"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        FR_Table.Delete
        FR_Table.MoveNext
    Loop
End If

FR_Table.Close
FR_Database.Close

cmdRefresh3_Click

End Sub

Private Sub cmdCode_Click(Index As Integer)

fraWS.Enabled = True

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)
 
Select Case CLng(Mid(lblCode(Index).Caption, 1, 3))
Case 400 To 410

Case Else
    Dim sSQL As String
    sSQL = "SELECT [CODE_ID] FROM [WORK SHEET] " & _
           "WHERE [DATE_ID]       =#" & DATE_ID & "# AND " & _
                 "[OP_ID]      =" & OP_ID & " AND " & _
                 "[MACHINE_ID] = " & MACHINE_ID & " AND " & _
                 "[TOTAL TIME] = 0 AND [CODE_ID] NOT IN (400,401,402,403,404,405,407) "
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
    If (FR_Table.RecordCount <> 0) Then
        MsgBox "A Work Order is not closed Select and Stop", vbInformation, "ATC Tracking System"
        FR_Table.Close
        FR_Database.Close
        Exit Sub
    End If
    FR_Table.Close
    FR_Database.Close

End Select

Select Case Index
Case 0 To 6         ' Code                                ,Description            ,Time
        Work_Codes CLng(Mid(lblCode(Index).Caption, 1, 3)), cmdCode(Index).Caption, 0, frmWorkSheet1
End Select

txtWorkOrder.SetFocus

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub



Private Sub cmdDelete_Click()

Dim iAns As Integer
iAns = MsgBox("Delete item " & WS_ID, vbQuestion + vbYesNo, "OEE Daily Work Sheet")

If (iAns = vbYes) Then

        Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)
 
        Dim sSQL As String
        sSQL = "SELECT * FROM [WORK SHEET] WHERE [WS_ID]=" & WS_ID
        
        Set FR_Table = FR_Database.OpenRecordset(sSQL)
         
        If (FR_Table.RecordCount <> 0) Then
            FR_Table.Delete
        End If
        FR_Table.Close
        FR_Database.Close
        
        WS_ID = -1
         
        cmdRefreshDisplay_Click
        MSFlexGrid1_Click
        
        DTPicker1.value = Format(Time, "hh:mm am/pm")
        DTPicker2.value = Format(Time, "hh:mm am/pm")
End If

End Sub

Private Sub cmdExit_Click()

If (WS_ID <> 0) Then
    'VALID_LASER_ID = 0
    '[1] DETERMINE HANDLER SIDE A/B

    
    '[2] DETERMINE LOGO MODE
    
    Dim ValidChar As String
    ValidChar = "N12379CHJ"
    
    Dim SearchChar As String
    SearchChar = Mid(ATC_PART_ID, 10, 1)
        
    If (Len(SearchChar) = 0) Then
        LOGO_MODE = LOGO_SIDE
    Else
        Dim MYPOS As Integer
        MYPOS = InStr(1, ValidChar, SearchChar, 1)
        Select Case MYPOS
        Case 0
            'NOT FOUND     MAG
                LOGO_MODE = LOGO_SIDE
        Case Else
            'FOUND     NON MAG
                LOGO_MODE = LOGO_ATC
        End Select
    End If
        
    If (PartLookup(ATC_PART_ID) = True) Then
    Else
        TEXT_ID = "NA"
    End If
        
    'Dim sSQL As String
    'Set FR_Database = OpenDatabase(ATC_LASER_BD)
    'sSQL = "SELECT * FROM [FIXTURE] WHERE [CAPTION] ='" & SERIES_ID & "'"
    'Set FR_Table = FR_Database.OpenRecordset(sSQL)
    'If (FR_Table.RecordCount <> 0) Then
    '    MATRIX_ID = FR_Table.Fields("[MATRIX ID]")
    '    CASE_ID = Mid(SERIES_ID, 4, 1)
    'Else
    '    MsgBox "No Table Available", vbCritical + vbInformation, "ATC"
    '    Exit Sub
    'End If
    
    Select Case CASE_ID
    Case "A"
      '      HANDLER_ID = 1
    Case "B"
     '       HANDLER_ID = 2
    End Select
        
    'VALID_LASER_ID = 1
End If

Unload Me
frmOPScreen.Show


End Sub

Private Sub cmdNext_Click()

DTPicker3.value = DateAdd("D", 1, DTPicker3.value)
DATE_ID = DTPicker3.value
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdPrevious_Click()

DTPicker3.value = DateAdd("D", -1, DTPicker3.value)
DATE_ID = DTPicker3.value
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh1_Click()

DATE_ID = DTPicker3.value
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdRefresh3_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT     [DEFECTS].[DF_ID]," & _
              "[DEFECT LIST].[DESCRIPTION]," & _
                  "[DEFECTS].[QTY] " & _
        "FROM [WORK SHEET],[DEFECT LIST],[DEFECTS] " & _
        "WHERE [WORK SHEET].[WS_ID]     = [DEFECTS].[WS_ID] AND " & _
             "[DEFECT LIST].[DEFECT_ID] = [DEFECTS].[DEFECT_ID] AND " & _
              "[WORK SHEET].[WS_ID]     = " & WS_ID & _
      " ORDER BY [DEFECT LIST].[DESCRIPTION]"

sSQLF = "    ||Description                           |>Quantity"
 
Data3.RecordSource = sSQL
Data3.Refresh

MSFlexGrid3.FormatString = sSQLF

End Sub

Private Sub cmdRefreshDisplay_Click()

'DATE_ID = Format(DATE_ID, "m/d/yyyy")

Dim sSQL As String
Dim sSQLF As String
 
Select Case 0
Case 0
        sSQL = "SELECT  [WS_ID]," & _
                       "[CODE_ID]," & _
                       "[WORK ORDER]," & _
                       "[ATC PART]," & _
                "format([START TIME],'h:mm AM/PM')," & _
                "format([STOP TIME],'h:mm AM/PM')," & _
                       "[TOTAL TIME],[QUANTITY]," & _
                       "[REJECTS],[RESTOCK] " & _
               "FROM [WORK SHEET] " & _
               "WHERE [DATE_ID]      =#" & DATE_ID & "# AND " & _
                     "[OP_ID]     =" & OP_ID & " AND " & _
                     "[MACHINE_ID]= " & MACHINE_ID & " " & _
               "ORDER BY [WS_ID] DESC"
Case 1

        sSQL = "SELECT [WS_ID],[CODE_ID],[WORK ORDER],[ATC PART]," & _
                       "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
                      "[TOTAL TIME],[QUANTITY],[REJECTS],[RESTOCK] " & _
               "FROM [WORK SHEET] " & _
               "WHERE [OP_ID]=" & OP_ID & " AND " & _
                     "[MACHINE_ID]= " & MACHINE_ID & " " & _
               "ORDER BY [WS_ID] DESC"
End Select


sSQLF = "    ||^Code|<W.O./Lot#                |<Atc Part            |^Start            |^Stop     "
sSQLF = sSQLF & "    |Time  |Quantity|>Rejects|>Restock"

Data1.RecordSource = sSQL
Data1.Refresh

MSFlexGrid1.FormatString = sSQLF

End Sub

Private Sub cmdReset_Click()

WS_ID = -1

DTPicker3.value = Date
DATE_ID = DTPicker3.value
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub cmdStopTime_Click()

DTPicker2.value = Format(Time, "hh:mm am/pm")

Dim sTime As String
If (DTPicker1.value > DTPicker2.value) Then
    sTime = DateDiff("n", DTPicker1.value, DTPicker2.value) + 1440
Else
    sTime = DateDiff("n", DTPicker1.value, DTPicker2.value)
End If

txtTotalTime.Text = sTime
 
Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET] WHERE [WS_ID]=" & WS_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    FR_Table.Edit
    FR_Table.Fields("[START TIME]") = Format(DTPicker1.value, "hh:mm am/pm")
    FR_Table.Fields("[STOP TIME]") = Format(DTPicker2.value, "hh:mm am/pm")
    FR_Table.Fields("[TOTAL TIME]") = Val(txtTotalTime.Text)
    FR_Table.Update
End If
FR_Table.Close
FR_Database.Close

cmdRefreshDisplay_Click

End Sub


Private Sub cmdUpdate_Click()

Data5.UpdateRecord

cmdRefresh3_Click

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT SUM([QTY]) AS SUM_REJECTS FROM [DEFECTS] " & _
       "WHERE [WS_ID]=" & WS_ID & " GROUP BY [WS_ID]"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    txtDefects.Text = FR_Table.Fields("[SUM_REJECTS]")
End If

FR_Table.Close
FR_Database.Close

Data4.UpdateRecord
cmdRefreshDisplay_Click


End Sub

Private Sub cmdUpdateRecord_Click()

Data4.UpdateRecord
cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub DTPicker2_LostFocus()

If (DTPicker1.value > DTPicker2.value) Then
    txtTotalTime.Text = DateDiff("n", DTPicker1.value, DTPicker2.value)
    txtTotalTime.Text = txtTotalTime.Text + 1440
Else
    txtTotalTime.Text = DateDiff("n", DTPicker1.value, DTPicker2.value)
End If

 
cmdRefreshDisplay_Click


End Sub


Private Sub Form_Activate()

' OOE EL,TR,LS

DTPicker1.value = Format(Time, "hh:mm am/pm")
DTPicker2.value = Format(Time, "hh:mm am/pm")

cmdRefreshDisplay_Click
MSFlexGrid1_Click

End Sub

Private Sub Form_Load()

Caption = "OEE Work Sheet          " & ATC_DWG & "      " & ATC_VERSION

lblDate2.Caption = Date
lblUser.Caption = strComputerName
lblIP.Caption = IP_ADDRESS
lblTime.Caption = Format(Time, "HH:MM AM/PM")

Data1.DatabaseName = DB_OEE_WORKSHEET
Data2.DatabaseName = DB_OEE_WORKSHEET
Data3.DatabaseName = DB_OEE_WORKSHEET
Data4.DatabaseName = DB_OEE_WORKSHEET
Data5.DatabaseName = DB_OEE_WORKSHEET
 
Data6.DatabaseName = ATC_LASER_BD
 
MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 11600
MSFlexGrid1.Height = 6000
MSFlexGrid1.ForeColorSel = vbBlack

MSFlexGrid2.Width = 3000
MSFlexGrid2.Height = 2400

MSFlexGrid3.Width = 3700
MSFlexGrid3.Height = 2400

lblNote.Caption = "Note : ATC Part Format chr 10" & vbNewLine & "Non Mag IN(N,1,3,5,7,9C,H,J)"
LabeMachine_ID.Caption = MACHINE_ID

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String

sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
   ' MACHINE_TYPE_ID = FR_Table.Fields("[TYPE NUMBER]")
  '  MACHINE_TYPE = FR_Table.Fields("[TYPE NUMBER]")
 '   MACHINE_NUMBER = FR_Table.Fields("[MACHINE]")
'    MACHINE_DESCRIPTION = FR_Table.Fields("[DESCRIPTION]")
End If


lblMachine.Caption = MACHINE_NUMBER & "   " & MACHINE_DESCRIPTION

'=============================================================================
'   LOAD DEFECT LIST
'=============================================================================

Select Case Mid(DEPT_ID, 1, 1)
Case "T"
        sSQL = "SELECT [DEFECT_ID],[DESCRIPTION]  " & _
                "FROM [DEFECT LIST] WHERE [TR]='T'"
Case "E"
        sSQL = "SELECT [DEFECT_ID],[DESCRIPTION]  " & _
                "FROM [DEFECT LIST] WHERE [ET]='E'"
Case "L"
        sSQL = "SELECT [DEFECT_ID],[DESCRIPTION]  " & _
                "FROM [DEFECT LIST] WHERE [LS]='L'"
End Select
 
Data2.RecordSource = sSQL
Data2.Refresh
 
Dim sSQLF As String
sSQLF = "    ||<Defect Description             "

MSFlexGrid2.FormatString = sSQLF
 
sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID]=" & OP_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    txtOperator.Caption = FR_Table.Fields("[FIRST]") & " " & FR_Table.Fields("[LAST]")
    txtShift.Caption = FR_Table.Fields("[SHIFT_ID]")
End If

   
'=============================================================================
'   LOAD DEPT CODES
'=============================================================================
   
Dim i As Integer
For i = 0 To 2
    cmdCode(i).Visible = True
    lblCode(i).Visible = True
Next i
  
sSQL = "SELECT [CODE_ID],[DESCRIPTION] " & _
       "FROM [CODES ET] " & _
       "WHERE [TYPE]=" & MACHINE_TYPE & " AND [ACTIVE] = 1"
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

i = 0
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        lblCode(i).Caption = FR_Table.Fields("[CODE_ID]")
        cmdCode(i).Caption = FR_Table.Fields("[DESCRIPTION]")
        lblCode(i).Visible = True
        cmdCode(i).Visible = True
        i = i + 1
        FR_Table.MoveNext
    Loop
End If

FR_Table.Close
FR_Database.Close
                
 
sSQL = "SELECT [TRAY_ID],[TITLE],[CASE]" & _
       "FROM [TBL TRAY CONFIG] " & _
       "ORDER BY [CASE],[TITLE]"
      
sSQLF = "     ||<Laser Tray Title                   |^Case"

Data6.RecordSource = sSQL
Data6.Refresh

MSFlexGrid6.FormatString = sSQLF
        
DTPicker3.value = Date$
DATE_ID = Date$
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Select Case UnloadMode
Case 0
        frmOPScreen.Show
Case 1
   
End Select

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then

End If

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
WS_ID = Val(MSFlexGrid1.Text)

If (WS_ID <> 0) Then
    fraWS.Enabled = True
Else
    fraWS.Enabled = False
End If

MSFlexGrid1.Col = 2
CODE_ID = Val(MSFlexGrid1.Text)

fraWS.Caption = "CODE_ID : " & CODE_ID

MSFlexGrid1.Col = 4
ATC_PART_ID = MSFlexGrid1.Text

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET] WHERE [WS_ID]=" & WS_ID
Data4.RecordSource = sSQL
Data4.Refresh
        
cmdRefresh3_Click
        
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10
               
End Sub

Private Sub MSFlexGrid2_Click()

MSFlexGrid2.Col = 1
DEFECT_ID = Val(MSFlexGrid2.Text)

MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1

End Sub

Private Sub MSFlexGrid3_Click()

fraQTY.Enabled = True

MSFlexGrid3.Col = 1
DF_ID = Val(MSFlexGrid3.Text)

Dim sSQL As String
sSQL = "SELECT * FROM [DEFECTS] WHERE [DF_ID]=" & DF_ID
Data5.RecordSource = sSQL
Data5.Refresh

MSFlexGrid3.Col = 0
MSFlexGrid3.ColSel = MSFlexGrid3.Cols - 1

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
    
    Select Case CODE_ID
    Case 660        'WORK ORDERS
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
    
    Dim ValidChar As String
    ValidChar = "N12379CHJ"
    
    Dim SearchChar As String
    SearchChar = Mid(ATC_PART_ID, 10, 1)
    
    Dim MYPOS As Integer
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

End Select

End Sub

Private Sub txtDefects_GotFocus()
txtDefects.SelStart = 0
txtDefects.SelLength = Len(txtDefects)
End Sub

Private Sub txtLot_LostFocus()
txtLot.Text = UCase(txtLot.Text)
End Sub

Private Sub txtOrderQty_GotFocus()
txtOrderQty.SelStart = 0
txtOrderQty.SelLength = Len(txtOrderQty)
End Sub

Private Sub txtQty_GotFocus()
txtQty.SelStart = 0
txtQty.SelLength = Len(txtQty)
End Sub

Private Sub txtRestock_GotFocus()
txtRestock.SelStart = 0
txtRestock.SelLength = Len(txtRestock)
End Sub

Private Sub txtTotalTime_GotFocus()
txtTotalTime.SelStart = 0
txtTotalTime.SelLength = Len(txtTotalTime)
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
 
Data4.UpdateRecord

cmdRefreshDisplay_Click

End Sub


Public Sub Work_Codes(iCode As Long, sDescription As String, iTotalTime As Integer, frmWorkSheet As Form)

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET] " & _
       "WHERE [DATE_ID]       =#" & DATE_ID & "# AND " & _
             "[OP_ID]      =" & OP_ID & " AND " & _
             "[MACHINE_ID] = " & MACHINE_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
FR_Table.AddNew
WS_ID = FR_Table.Fields("WS_ID")

FR_Table.Fields("[OP_ID]") = OP_ID
FR_Table.Fields("[DATE_ID]") = DATE_ID
FR_Table.Fields("[MACHINE_ID]") = MACHINE_ID
FR_Table.Fields("[CODE_ID]") = iCode

Select Case sDescription
Case "Planned Downtime"
        FR_Table.Fields("[WORK ORDER]") = "Planned DT"
Case "Unplanned Downtime"
        FR_Table.Fields("[WORK ORDER]") = "Unplanned DT"
Case Else
        If (Len(sDescription) > 14) Then
            sDescription = Mid(sDescription, 1, 14)
        End If
        FR_Table.Fields("[WORK ORDER]") = sDescription
End Select

FR_Table.Fields("[RESTOCK]") = 0
FR_Table.Fields("[START TIME]") = Format(Time, "hh:mm am/pm")
FR_Table.Fields("[TOTAL TIME]") = iTotalTime

FR_Table.Update

FR_Table.Close
FR_Database.Close

End Sub

Public Function PartLookup(sPart As String) As Boolean

Select Case Mid$(sPart, 5, 1)
Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
Case Else
             MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
             PartLookup = False
             Exit Function
End Select

Select Case Mid$(sPart, 6, 1)
Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "R"
Case Else
             MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
             PartLookup = False
             Exit Function
End Select

Select Case Mid$(sPart, 7, 1)
Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
Case Else
             MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
             PartLookup = False
             Exit Function
End Select

Select Case Mid$(sPart, 8, 1)
Case "A", "B", "C", "D", "F", "G", "J", "K", "L", "M", "N"
Case Else
             MsgBox "Cap Tolerance Not Valid", vbInformation, "ATC Part Number"
             PartLookup = False
             Exit Function
End Select

SERIES_ID = Mid(sPart, 1, 4)
            
TEXT_ID = Mid(sPart, 5, 4)
      
PartLookup = True
            
End Function
