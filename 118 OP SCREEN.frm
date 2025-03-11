VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOPScreen 
   Caption         =   "118  DPSS Tray Production Screen"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16560
   Icon            =   "118 OP SCREEN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   16560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTR 
      Interval        =   1000
      Left            =   10680
      Top             =   1320
   End
   Begin VB.CommandButton CommandParaylene 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Parylene Demasking"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3420
      Width           =   3400
   End
   Begin VB.CommandButton CommandLimits 
      Caption         =   "Find Limits"
      Height          =   250
      Left            =   10440
      TabIndex        =   67
      Top             =   5640
      Width           =   1800
   End
   Begin VB.CommandButton CommandLoad 
      Caption         =   "Load LaserPro"
      Height          =   250
      Left            =   10440
      TabIndex        =   66
      Top             =   5280
      Width           =   1800
   End
   Begin VB.CommandButton CommandSystem 
      Caption         =   "System Initialize"
      Height          =   250
      Left            =   10440
      TabIndex        =   65
      Top             =   4920
      Width           =   1800
   End
   Begin VB.CommandButton CommandInitialize_Controller 
      Caption         =   "Abort Motion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   10440
      TabIndex        =   64
      Top             =   6000
      Width           =   1800
   End
   Begin VB.CommandButton CommandMain 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tray Laser Main"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4080
      Width           =   3400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Power Factors"
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
      Left            =   12960
      TabIndex        =   52
      Top             =   6240
      Width           =   3400
   End
   Begin VB.CommandButton CommandTray 
      Caption         =   "Tray  Configuration"
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
      Left            =   12960
      TabIndex        =   51
      Top             =   6660
      Width           =   3400
   End
   Begin VB.CommandButton CommandMotion 
      Caption         =   "Motion"
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
      Left            =   12960
      TabIndex        =   50
      Top             =   5820
      Width           =   3400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DB Backup"
      Enabled         =   0   'False
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8760
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.CommandButton CommandUpdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update WO Schedule"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4500
      Width           =   3400
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   11760
      MaxLength       =   1
      PasswordChar    =   "*"
      TabIndex        =   43
      Text            =   "X"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton CommandBackUpDB 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Back Up DB"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8400
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.CommandButton CommandWorkSheetReview 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Work Sheet Review"
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
      Height          =   420
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3000
      Width           =   3400
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
      Left            =   12960
      TabIndex        =   38
      Top             =   960
      Width           =   3375
      Begin VB.CommandButton CommandT 
         Caption         =   "Test"
         Height          =   300
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton CommandP 
         Caption         =   "Prod"
         Height          =   300
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   11520
      PasswordChar    =   "*"
      TabIndex        =   37
      Text            =   "XXXXX"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox TextLOCATION 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   12840
      TabIndex        =   35
      Text            =   "LS"
      Top             =   7920
      Width           =   495
   End
   Begin VB.Data Data2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data2 FROM [FIXTURE]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\111 MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8760
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Frame fraWS 
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
      Height          =   6975
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "CAPTION"
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
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   6120
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "CASE SIZE"
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
         Left            =   3840
         TabIndex        =   31
         Text            =   "Z"
         Top             =   6120
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   29
         Top             =   5040
         Width           =   2505
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   27
         Top             =   4560
         Width           =   1035
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
         TabIndex        =   15
         Text            =   "12345"
         ToolTipText     =   "Start Qty [START QTY]"
         Top             =   2280
         Width           =   825
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
         TabIndex        =   14
         Text            =   "1234567890"
         Top             =   1800
         Width           =   2280
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
         TabIndex        =   13
         Text            =   "123456789012"
         Top             =   840
         Width           =   2280
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
         TabIndex        =   12
         Text            =   "12345"
         ToolTipText     =   "QUANTITY"
         Top             =   2760
         Width           =   825
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
         Left            =   3600
         TabIndex        =   11
         Text            =   "12345"
         ToolTipText     =   "REJECTS"
         Top             =   2280
         Width           =   825
      End
      Begin VB.TextBox txtTotalTime 
         DataField       =   "TOTAL TIME"
         DataSource      =   "Data4"
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
         TabIndex        =   10
         Text            =   "0"
         ToolTipText     =   "Total time"
         Top             =   3960
         Width           =   600
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
         TabIndex        =   9
         ToolTipText     =   "Test for Valid Tolerance on exit field"
         Top             =   1320
         Width           =   2280
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
         Left            =   3600
         TabIndex        =   8
         Text            =   "12345"
         ToolTipText     =   "Restock"
         Top             =   2760
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "START TIME"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
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
         Format          =   52232194
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         DataField       =   "DATE"
         DataSource      =   "Data4"
         Height          =   375
         Left            =   1560
         TabIndex        =   25
         Top             =   360
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
         Format          =   52232193
         CurrentDate     =   38117
      End
      Begin VB.Label Label13 
         Caption         =   "Matrix_ID:"
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
         Left            =   240
         TabIndex        =   34
         Top             =   6120
         Width           =   1035
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         DataField       =   "MATRIX ID"
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
         Left            =   1320
         TabIndex        =   33
         Top             =   6120
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Logo"
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
         TabIndex        =   30
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Mark Text "
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
         TabIndex        =   28
         Top             =   4560
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "Date_ID:"
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
         TabIndex        =   26
         Top             =   360
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
         TabIndex        =   24
         Top             =   2280
         Width           =   1035
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
         TabIndex        =   23
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "Total Time (m):"
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
         Left            =   480
         TabIndex        =   22
         Top             =   3960
         Width           =   1665
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
         TabIndex        =   21
         Top             =   840
         Width           =   1305
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
         TabIndex        =   20
         Top             =   2760
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "Defects :"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   2280
         Width           =   1035
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
         TabIndex        =   18
         Top             =   1320
         Width           =   1275
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
         Left            =   2520
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.CommandButton CommandConfiguration 
      Caption         =   "Configuration"
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
      Left            =   12960
      TabIndex        =   6
      Top             =   5400
      Width           =   3400
   End
   Begin VB.CommandButton CommandRun 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Run Screen"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2580
      Width           =   3400
   End
   Begin VB.CommandButton cmdWS 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Work Sheet"
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
      Height          =   420
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   3400
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 FROM [BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Frame Frame7 
      Caption         =   " Operator Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton cmdOperator 
         Caption         =   "Operator"
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "118 OP SCREEN.frx":0CCA
         Height          =   1215
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2143
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
      End
      Begin VB.Label lblOperator 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2835
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TimerExitProgram_Timer"
      Height          =   300
      Left            =   10680
      TabIndex        =   69
      Top             =   960
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label LabelINITIALIZE_TRAY 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXX"
      Height          =   300
      Left            =   11880
      TabIndex        =   63
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label7 
      Caption         =   "INITIALIZE_TRAY"
      Height          =   300
      Left            =   10320
      TabIndex        =   62
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Label5 
      Caption         =   "MACHINE_ID"
      Height          =   300
      Left            =   10320
      TabIndex        =   61
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label Label8 
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
      Left            =   10320
      TabIndex        =   60
      Top             =   3960
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
      Left            =   11880
      TabIndex        =   59
      Top             =   3960
      Width           =   495
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
      Height          =   360
      Left            =   10320
      TabIndex        =   58
      Top             =   3480
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
      Left            =   11880
      TabIndex        =   57
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label4 
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
      Left            =   10320
      TabIndex        =   56
      Top             =   3000
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
      Left            =   11880
      TabIndex        =   55
      Top             =   3000
      Width           =   495
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
      Left            =   4440
      TabIndex        =   53
      Top             =   7320
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "118 LASER MATRIX.mdb"
      Height          =   300
      Left            =   11760
      TabIndex        =   49
      Top             =   8760
      Width           =   4515
   End
   Begin VB.Label LabelMACHINE_ID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXX"
      Height          =   300
      Left            =   11880
      TabIndex        =   47
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label LabelDB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DB"
      Height          =   300
      Left            =   11760
      TabIndex        =   46
      Top             =   8400
      Width           =   4515
   End
   Begin VB.Label LabelWODate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      Height          =   300
      Left            =   14640
      TabIndex        =   45
      Top             =   7920
      Width           =   1635
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   12360
      Picture         =   "118 OP SCREEN.frx":0CDE
      Top             =   120
      Width           =   4170
   End
   Begin VB.Label LabelDBMode 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DB Mode"
      Height          =   300
      Left            =   13440
      TabIndex        =   36
      Top             =   7920
      Width           =   1035
   End
End
Attribute VB_Name = "frmOPScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOperator_Click()

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [OP_ID],[FIRST]& ' ' & [LAST] FROM [BARCODE] "

sSQL = sSQL & "WHERE mid([DEPT_ID],1,1)='L' AND [ACTIVE]=1 ORDER BY [LAST],[OP_ID]"
  
sSQLF = "    ||<Operator                         "

Data1.RecordSource = sSQL
Data1.Refresh
 
MSFlexGrid1.FormatString = sSQLF
MSFlexGrid1.Width = 3400
MSFlexGrid1.Height = 7400

End Sub

Private Sub cmdWS_Click()

If (OP_ID = 0) Then
    MsgBox "No Operator Selected", vbInformation, "Daily OEE"
    Exit Sub
End If

If (MACHINE_ID = 0) Then
    MsgBox "No Equipment Selected", vbInformation, "Daily OEE"
    Exit Sub
End If

Dim SHIFT_TIME As String
SHIFT_TIME = "3 AM"

Dim START_TIME As Date
START_TIME = Format(Time, "h AM/PM")

If (START_TIME < SHIFT_TIME) Then
    'Change Date -1
    DATE_ID = DateAdd("d", -1, Date$)
Else
    DATE_ID = Date$
End If

'frmMain.Hide

frmOPScreen.Hide

frmWorkSheet1.Show

End Sub

Private Sub Command1_Click()

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

Private Sub Command2_Click()

'CASE_ID = "C"
'TRAY_ID = 2
'SIZE_LOC_ID = 2
'POWER_ID = 27

If TRAY_ID = 0 Then Exit Sub

frmPowerFactors.Show

End Sub


Private Sub CommandBackUpDB_Click()

Dim Message, Title, Default, MyValue

Message = "Enter a USB Drive Letter B to G"   ' Set prompt.
Title = "InputBox Drive for Data Base Back Up"   ' Set title.
Default = "E"   ' Set default.

' Display message, title, and default value.
MyValue = InputBox(Message, Title, Default)

If MyValue = "" Then Exit Sub

On Error GoTo Network_Mode_ErrorAll

Dim dTime As Single
Dim SourceFile As String
Dim DestinationFile As String

Dim FSO As New FileSystemObject

Screen.MousePointer = vbHourglass

Dim i As Integer
Dim Seconds As Long
Dim MinSec As String
  
 
SourceFile = "C:\ATC\118 LASER MATRIX.MDB"
DestinationFile = MyValue & ":\118 LASER MATRIX.MDB"

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

Private Sub CommandConfiguration_Click()

frmConfiguration.Show

End Sub

Private Sub CommandInitialize_Controller_Click()
Initialize_Controller
End Sub

Private Sub CommandLimits_Click()

Screen.MousePointer = vbHourglass

CommandSystem.Caption = "FindReverseLimit Z"
FindReverseLimit 3

CommandSystem.Caption = "FindReverseLimit X"
FindReverseLimit 1

CommandSystem.Caption = "FindReverseLimit Y"
FindReverseLimit 2

Screen.MousePointer = vbDefault

Dim sBuff As String
sBuff = "NI Motion Parameters Complete"

MsgBox sBuff, vbInformation, "ATC Laser Tray System"


CommandLimits.FontBold = True

End Sub

Private Sub CommandLoad_Click()

If (Load_Job = 0) Then
       Load_Job_From_File
End If

CommandLoad.FontBold = True

End Sub

Private Sub CommandMain_Click()

LOGO_MODE = LOGO_TOP
frmMain.Show

End Sub

Private Sub CommandMotion_Click()
frmMotion.Show      'NI Motion Functions
End Sub

Private Sub CommandP_Click()
 
Form_Activate
End Sub

Private Sub CommandParaylene_Click()

TRAY_ID = 15
POWER_ID = 15
SIZE_LOC_ID = 10

LASER_TXT1 = "ABCD"
LASER_TXT2 = "ABCD"
LASER_TXT3 = "ABCD"
LASER_TXT4 = "ABCD"

'frmDemaskingH.Show         'ATC 301-476           'H' Case Carrier Tray

TRAY_ID = 1
POWER_ID = 80
SIZE_LOC_ID = 10            'CASE SIZE H ABRASIVE PARAMETERS
LOGO_MODE = LOGO_ABRASIVE

frm413.Show

End Sub

Private Sub CommandRun_Click()

Static bInitialized As Boolean

LOGO_MODE = LOGO_TOP

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

If (TRAY_ID = 0) Then
    MsgBox "No Work Sheet has been set up or Fixture selected", vbInformation + vbCritical, "ATC Tray Laser"
    Exit Sub
End If

frmOPScreen.Hide

Initialize_Fire_Matrix

If (Load_Job = 0) Then
       Load_Job_From_File
End If

Select Case TRAY_ID
Case 13
            frm103.Show
Case 1
            frm413.Show
Case 2, 3, 4, 12
            frm414.Show
Case 5, 6, 7, 10, 11
            frm412.Show
Case 8
            frm10x10.Show
Case 9
            frm20x20.Show
End Select

End Sub

Private Sub CommandSystem_Click()

Screen.MousePointer = vbHourglass

Initialize_Controller

DisableHome

Load_Parameters

CommandSystem.Caption = "Load_Parameters"

Screen.MousePointer = vbDefault

CommandSystem.Caption = "System Initialize"
CommandSystem.FontBold = True

Dim sBuff As String

sBuff = "NI Motion Parameters Complete"

MsgBox sBuff, vbInformation, "ATC Laser Tray System"

End Sub

Private Sub CommandT_Click()
 
Form_Activate
End Sub

Private Sub CommandTray_Click()
 frmTray.Show        'Tray Configuration
End Sub

Private Sub CommandUpdate_Click()

On Error GoTo UpdateDBSchedule_Error
 
Dim SourceFile As String
Dim DestinationFile As String

Select Case LOCATION_ID
Case "JR"
        SourceFile = SERVER_DB_JR & "WO SCHED MASTER.MDB"
Case "NY"
        SourceFile = SERVER_DB_NY & "WO SCHED MASTER.MDB"
      '  Exit Sub
End Select

Screen.MousePointer = vbHourglass

DestinationFile = "C:\ATC\WO SCHED MASTER.MDB"

Dim FSO As New FileSystemObject

FSO.CopyFile SourceFile, DestinationFile, True

Screen.MousePointer = vbDefault

LabelWODate.Caption = Format(Date, "MM/DD/YY") & " " & Format(Time, "hh:mm AM/PM")

Exit Sub
UpdateDBSchedule_Error:

Screen.MousePointer = vbDefault

LabelWODate.Caption = "Unsuccesful"

'MsgBox "Unsuccesful", vbCritical, "ATC DataBase System"

End Sub

Private Sub CommandWorkSheetReview_Click()

frmReviewWS.Show

End Sub

Private Sub Form_Activate()

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

LabelDB.Caption = DB_OEE_WORKSHEET

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET] WHERE [WS_ID]=" & WS_ID

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then

    If IsNull(FR_Table.Fields("[WORK ORDER]")) = False Then
        txtWorkOrder.Text = FR_Table.Fields("[WORK ORDER]")
    Else
        txtWorkOrder.Text = ""
    End If
    
    If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
        txtATCPart.Text = FR_Table.Fields("[ATC PART]")
    Else
        txtATCPart.Text = ""
    End If
    
    If IsNull(FR_Table.Fields("[LOT NUM]")) = False Then
        txtLot.Text = FR_Table.Fields("[LOT NUM]")
    Else
        txtLot.Text = ""
    End If
        
    txtOrderQty.Text = FR_Table.Fields("[QUANTITY]")
    txtSQ.Text = FR_Table.Fields("[START QTY]")
    txtDefects.Text = FR_Table.Fields("[REJECTS]")
    txtRestock.Text = FR_Table.Fields("[RESTOCK]")
         
    If IsNull(FR_Table.Fields("[START TIME]")) = False Then
        DTPicker1.value = FR_Table.Fields("[START TIME]") ' Format(Date, "MM/dd/yyyy ") & Format(Time, "hh:mm am/pm")
    Else
        DTPicker1.value = Format(Time, "hh:mm am/pm")
    End If
    
    DTPicker3.value = FR_Table.Fields("[DATE_ID]")
    txtTotalTime.Text = FR_Table.Fields("[TOTAL TIME]")
Else
    txtWorkOrder.Text = ""
    txtATCPart.Text = ""
    txtLot.Text = ""
    txtOrderQty.Text = ""
    txtSQ.Text = ""
    txtDefects.Text = ""
    txtRestock.Text = ""
    DTPicker1.value = Format(Time, "hh:mm am/pm")
   
    DTPicker3.value = Date
    txtTotalTime.Text = ""
End If
FR_Table.Close
FR_Database.Close


Text1.Text = TEXT_ID

LASER_TXT1 = TEXT_ID

Select Case LOGO_MODE
Case LOGO_ATC
        Text2.Text = "Non Mag Logo Top ATC"
Case LOGO_SIDE
        Text2.Text = "Mag Side Logo"
End Select

frmOPScreen.Top = FORM_LOC_Y
frmOPScreen.Left = FORM_LOC_X

LabelSIZE_LOC_ID.Caption = SIZE_LOC_ID
LabelTRAY_ID.Caption = TRAY_ID
LabelPOWER_ID.Caption = POWER_ID

LabelINITIALIZE_TRAY.Caption = INITIALIZE_TRAY

End Sub


Private Sub Form_Load()

Caption = "DPSS Tray Laser Production Screen" & Space(8) & ATC_DWG & Space(8) & ATC_VERSION
Caption = Caption & Space(8) & "IP " & IP_ADDRESS & Space(8) & strComputerName

Data1.DatabaseName = DB_ELECT_OP_MACHINE
Data2.DatabaseName = ATC_LASER_BD

LabelMACHINE_ID.Caption = MACHINE_ID

TextLOCATION.Text = LOCATION_ID

DTPicker3.value = Date

CommandUpdate_Click

cmdOperator_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim iAns As Integer
iAns = MsgBox("Exit Program", vbYesNo, "ATC Laser System")
If (iAns = vbYes) Then

    ConfigComputer_DB (1)
    
    FORM_LOC_Y = frmOPScreen.Top
    FORM_LOC_X = frmOPScreen.Left
    
    Configuration (FWRITE)
            
    Cancel = 0
    
    End
Else
    Cancel = 1
End If

End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sBuff As String

sBuff = UCase(txtPassword.Text)

If (Button = 2 And Shift = 1) Then
       
    Select Case sBuff
    Case "MIKE" & Mid(Format(Date, "ddd"), 1, 1), "ERIK" & Mid(Format(Date, "ddd"), 1, 1)
                fraPTMode.Enabled = True
    Case Else
             
    End Select
    
    txtPassword.Text = "XXXX"
   ' CommandBackUpDB.Visible = True
   CommandWorkSheetReview.Enabled = True
Else
    'CommandBackUpDB.Visible = False
    CommandWorkSheetReview.Enabled = False
    fraPTMode.Enabled = False
End If

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
OP_ID = Val(MSFlexGrid1.Text)

cmdWS.Enabled = True

MSFlexGrid1.Col = 2
lblOperator.Caption = MSFlexGrid1.Text
 
MSFlexGrid1.Col = 0
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '10
 
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)

If (KeyCode = 84 And Shift = 2) Then
    Beep
End If

End Sub

Private Sub tmrTR_Timer()

If Format(Time, "hh AM/PM") = "01 AM" Then
    ConfigComputer_DB (2)
    End
End If

Strangelove

End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub
