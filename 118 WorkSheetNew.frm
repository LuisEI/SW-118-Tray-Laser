VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmWorkSheet1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "118 Laser OEE Work Sheet"
   ClientHeight    =   12180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12180
   ScaleWidth      =   18270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Scan Work Order "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4920
      TabIndex        =   75
      Top             =   6840
      Width           =   11415
      Begin VB.CommandButton CommandAdd 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add Record"
         Height          =   300
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton CommandSet 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Select Fixture"
         Height          =   300
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1200
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   77
         Text            =   "100E102FQX"
         ToolTipText     =   " "
         Top             =   720
         Width           =   1680
      End
      Begin VB.TextBox TextWorkOrder 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   76
         Text            =   "789934060006"
         ToolTipText     =   " "
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label labelPOWER_ID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10560
         TabIndex        =   92
         Top             =   840
         Width           =   555
      End
      Begin VB.Label lblInfo 
         Caption         =   "POWER_ID"
         Height          =   300
         Index           =   16
         Left            =   9480
         TabIndex        =   91
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblTRAY_ID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   8760
         TabIndex        =   90
         Top             =   840
         Width           =   555
      End
      Begin VB.Label lblInfo 
         Caption         =   "TRAY_ID"
         Height          =   300
         Index           =   14
         Left            =   7920
         TabIndex        =   89
         Top             =   840
         Width           =   795
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
         TabIndex        =   88
         Top             =   360
         Width           =   720
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
         Index           =   13
         Left            =   9120
         TabIndex        =   87
         Top             =   360
         Width           =   1035
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
         Left            =   7320
         TabIndex        =   86
         Top             =   360
         Width           =   675
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
         Left            =   5280
         TabIndex        =   85
         Top             =   360
         Width           =   1155
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
         TabIndex        =   84
         Top             =   360
         Width           =   840
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
         Left            =   6480
         TabIndex        =   83
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblCode_ID 
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
         Left            =   4440
         TabIndex        =   82
         Top             =   360
         Width           =   600
      End
      Begin VB.Label LabelTRAY_ID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3000
         TabIndex        =   81
         Top             =   720
         Width           =   4665
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
         Left            =   120
         TabIndex        =   80
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "W.O./Lot#:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   79
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdRefresh2 
      Caption         =   "Refresh2"
      Height          =   300
      Left            =   240
      TabIndex        =   69
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data7 
      Caption         =   "Data7 FROM [TBL Power]"
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
      Top             =   3600
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fixture "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   53
      Top             =   840
      Width           =   4695
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
         Index           =   5
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2880
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.OptionButton Option4 
         Caption         =   "H Case"
         Height          =   300
         Left            =   3480
         TabIndex        =   72
         Top             =   3360
         Width           =   900
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   3600
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
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   960
         Visible         =   0   'False
         Width           =   3600
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
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1440
         Visible         =   0   'False
         Width           =   3600
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
         Index           =   3
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1920
         Visible         =   0   'False
         Width           =   3600
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
         Index           =   4
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2400
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "E Case"
         Height          =   300
         Left            =   240
         TabIndex        =   57
         Top             =   3360
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton Option2 
         Caption         =   "C Case"
         Height          =   300
         Left            =   1320
         TabIndex        =   56
         Top             =   3360
         Width           =   900
      End
      Begin VB.OptionButton Option3 
         Caption         =   "B Case"
         Height          =   300
         Left            =   2400
         TabIndex        =   55
         Top             =   3360
         Width           =   900
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   15
         Left            =   3720
         TabIndex        =   54
         Top             =   480
         Width           =   735
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
         Index           =   5
         Left            =   120
         TabIndex        =   74
         Top             =   2880
         Visible         =   0   'False
         Width           =   600
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
         Left            =   120
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   600
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
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Visible         =   0   'False
         Width           =   600
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
         Index           =   2
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Visible         =   0   'False
         Width           =   600
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
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   1920
         Visible         =   0   'False
         Width           =   600
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
         Index           =   4
         Left            =   120
         TabIndex        =   63
         Top             =   2400
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 FROM [TBL SIZE LOC]"
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   48
      Text            =   "XXXX"
      Top             =   11520
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
      Height          =   3975
      Left            =   120
      TabIndex        =   22
      Top             =   4680
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
         TabIndex        =   33
         Top             =   2400
         Width           =   1875
      End
      Begin VB.TextBox txtRestock 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   32
         Text            =   "12345"
         ToolTipText     =   "Restock"
         Top             =   1920
         Width           =   825
      End
      Begin VB.CommandButton cmdStopTime 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Stop Time"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3360
         Width           =   1440
      End
      Begin VB.TextBox txtATCPart 
         BackColor       =   &H00FFFFC0&
         DataField       =   "ATC PART"
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   30
         ToolTipText     =   "Test for Valid Tolerance on exit field"
         Top             =   720
         Width           =   2280
      End
      Begin VB.TextBox txtTotalTime 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   29
         Text            =   "0"
         ToolTipText     =   "Total time"
         Top             =   3360
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
         TabIndex        =   28
         Top             =   2400
         Width           =   1875
      End
      Begin VB.TextBox txtDefects 
         BackColor       =   &H00C0FFC0&
         DataField       =   "REJECTS"
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
         Left            =   3480
         TabIndex        =   27
         Text            =   "12345"
         ToolTipText     =   "Defects"
         Top             =   1560
         Width           =   825
      End
      Begin VB.TextBox txtOrderQty 
         BackColor       =   &H00FFFFC0&
         DataField       =   "QUANTITY"
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
         Left            =   1560
         TabIndex        =   26
         Text            =   "12345"
         ToolTipText     =   "Units Produced"
         Top             =   1920
         Width           =   825
      End
      Begin VB.TextBox txtWorkOrder 
         BackColor       =   &H00FFFFC0&
         DataField       =   "WORK ORDER"
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   25
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   24
         Text            =   "1234567890"
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox txtSQ 
         BackColor       =   &H00FFFFC0&
         DataField       =   "START QTY"
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
         Left            =   1560
         TabIndex        =   23
         Text            =   "12345"
         ToolTipText     =   "Start Qty [START QTY]"
         Top             =   1560
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "START TIME"
         DataSource      =   "Data4"
         Height          =   360
         Left            =   360
         TabIndex        =   42
         Top             =   2880
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "h:mm AM/PM"
         Format          =   96731138
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "STOP TIME"
         DataSource      =   "Data4"
         Height          =   360
         Left            =   2400
         TabIndex        =   43
         Top             =   2880
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "h:mm AM/PM"
         Format          =   96731138
         CurrentDate     =   38117
      End
      Begin VB.Label lblInfo 
         Caption         =   "Restock :"
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
         Left            =   2520
         TabIndex        =   41
         Top             =   1920
         Width           =   975
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
         Index           =   8
         Left            =   240
         TabIndex        =   40
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         Caption         =   "Defects :"
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
         Index           =   6
         Left            =   2520
         TabIndex        =   39
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblInfo 
         Caption         =   "Quantity:"
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
         Left            =   240
         TabIndex        =   38
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label lblInfo 
         Caption         =   "W.O./Lot#:"
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
         Left            =   240
         TabIndex        =   37
         Top             =   360
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
         Left            =   360
         TabIndex        =   36
         Top             =   3360
         Width           =   1665
      End
      Begin VB.Label lblInfo 
         Caption         =   "Lot Number:"
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
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         Caption         =   "Start Qty:"
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
         Left            =   240
         TabIndex        =   34
         Top             =   1560
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
      Left            =   8160
      TabIndex        =   15
      Top             =   11160
      Width           =   5655
      Begin VB.CommandButton cmdRefresh1 
         Caption         =   "Refresh"
         Height          =   300
         Left            =   4380
         TabIndex        =   19
         Top             =   360
         Width           =   900
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Day  <<"
         Height          =   300
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   900
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Day  >>"
         Height          =   300
         Left            =   2580
         TabIndex        =   17
         Top             =   360
         Width           =   900
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   300
         Left            =   3480
         TabIndex        =   16
         Top             =   360
         Width           =   900
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1305
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
         Format          =   98369537
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
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.CommandButton cmdAddDefect 
      Caption         =   "Add  Defect to List >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   14
      Top             =   8520
      Width           =   2280
   End
   Begin VB.Frame fraQTY 
      Enabled         =   0   'False
      Height          =   735
      Left            =   8160
      TabIndex        =   10
      Top             =   10200
      Width           =   3015
      Begin VB.CommandButton cmdClear 
         Caption         =   "CLR Qty=0"
         Height          =   300
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Refresh"
         Height          =   300
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         DataField       =   "QTY"
         DataSource      =   "Data5"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Text            =   "0"
         ToolTipText     =   "Defects"
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdRefresh3 
      Caption         =   "<< Refresh3"
      Height          =   300
      Left            =   8280
      TabIndex        =   7
      Top             =   9720
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 FROM [WORK SHEET],[DEFECT LIST ET],[DEFECTS]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   5460
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [Defect List ET]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   4140
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   975
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1720
      _Version        =   393216
      BackColorSel    =   16776960
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   " FROM [WORK SHEET]"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   11640
      Width           =   2280
   End
   Begin VB.Data Data4 
      Caption         =   "Data4 FROM [WORK SHEET]"
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
      RecordSource    =   "WORK SHEET"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.CommandButton cmdRefreshDisplay 
      Caption         =   "Refresh Display"
      Height          =   300
      Left            =   5280
      TabIndex        =   1
      Top             =   2160
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
      Top             =   1320
      Visible         =   0   'False
      Width           =   3900
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   735
      Left            =   5400
      TabIndex        =   6
      Top             =   8880
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
      Height          =   735
      Left            =   7800
      TabIndex        =   8
      Top             =   8880
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
      Height          =   975
      Left            =   1560
      TabIndex        =   49
      Top             =   8760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1720
      _Version        =   393216
      BackColorSel    =   16776960
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
      FormatString    =   " FROM [TBL SIZE LOC]"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
      Height          =   975
      Left            =   4920
      TabIndex        =   68
      Top             =   4080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      _Version        =   393216
      BackColorSel    =   16776960
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   " FROM [TBL Power]"
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
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2700
      Left            =   12600
      Top             =   8520
      Width           =   5265
   End
   Begin VB.Label Label1 
      Caption         =   "Non Mag : ATC Logo (top)"
      Height          =   300
      Left            =   360
      TabIndex        =   71
      Top             =   10140
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Mag : Top Logo"
      Height          =   300
      Left            =   360
      TabIndex        =   70
      Top             =   10440
      Width           =   2235
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Extened E ATC Part Pos 10 7,8,B"
      Height          =   300
      Left            =   360
      TabIndex        =   52
      Top             =   11040
      Width           =   4155
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Extened E ATC Part Pos 9,10 E,E1,E2,EN,J,JN"
      Height          =   300
      Left            =   360
      TabIndex        =   51
      Top             =   10740
      Width           =   4155
   End
   Begin VB.Label LabelLogo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2760
      TabIndex        =   50
      Top             =   10200
      Width           =   2265
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   14280
      TabIndex        =   47
      Top             =   11760
      Width           =   1605
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default: "
      Height          =   300
      Left            =   16080
      TabIndex        =   46
      Top             =   11400
      Width           =   1605
   End
   Begin VB.Label lblDate2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      Height          =   300
      Left            =   14280
      TabIndex        =   45
      Top             =   11400
      Width           =   1605
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
      Height          =   300
      Left            =   16080
      TabIndex        =   44
      Top             =   11760
      Width           =   1605
   End
   Begin VB.Label lblNote 
      Caption         =   "Note : ATC Part "
      Height          =   300
      Left            =   360
      TabIndex        =   21
      Top             =   9840
      Width           =   4875
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   120
      Top             =   11400
      Width           =   4170
   End
   Begin VB.Label lblInfo 
      Caption         =   "Defect List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7800
      TabIndex        =   9
      Top             =   8520
      Width           =   1905
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
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   3945
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
      Left            =   3120
      TabIndex        =   3
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2745
   End
End
Attribute VB_Name = "frmWorkSheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
    
    VALID_LASER_ID = 0
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
        
    CODE_ID = CODE_ID - 700
    
    Dim sSQL As String
   ' Set FR_Database = OpenDatabase(ATC_LASER_BD)
   ' sSQL = "SELECT * FROM [FIXTURE] WHERE [MATRIX ID] =" & CODE_ID
   ' Set FR_Table = FR_Database.OpenRecordset(sSQL)
   ' If (FR_Table.RecordCount <> 0) Then
   '     MATRIX_ID = FR_Table.Fields("[MATRIX ID]")
  '     CASE_ID = Mid(SERIES_ID, 4, 1)
   ' Else
   '     MsgBox "No Table Available", vbCritical + vbInformation, "ATC"
    '    Exit Sub
   ' End If
    
    Select Case CASE_ID
    Case "A"
            
    Case "B"
            
    Case "C"
            
    Case "E"
                
    End Select
        
        
    '115 Keyence Laser 09/26/2013 CASE_ID "X" Extended E
    
    Select Case CASE_ID
    Case "E"
                Select Case Mid(ATC_PART_ID, 9, 1)
                Case "E", "J"
                                CASE_ID = "X"
                End Select
                Select Case Mid(ATC_PART_ID, 9, 2)
                Case "E1", "E2", "EN", "JN"
                                CASE_ID = "X"
                End Select
                Select Case Mid(ATC_PART_ID, 10, 1)
                Case "7", "8", "B"
                                CASE_ID = "X"
                End Select
    End Select
    
    VALID_LASER_ID = 1

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

Private Sub cmdRefresh2_Click()

Dim sSQL As String
Dim sSQLF As String
     
sSQL = "SELECT [SIZE_LOC_ID],[CASE NAME]," & _
              "format([FONT HEIGHT],'0.000') " & _
       "FROM [TBL SIZE LOC] WHERE [ACTIVE] = Yes AND [CASE]='" & CASE_ID & "'"
       
                                   
sSQLF = "   ||<Case Size           |Font Height "

Data6.RecordSource = sSQL
Data6.Refresh
 
MSFlexGrid6.FormatString = sSQLF

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
              "[WORK SHEET].[WS_ID]     = " & WS_ID & " " & _
       "ORDER BY [DEFECT LIST].[DESCRIPTION]"

sSQLF = "    ||<Defect Description|>Quantity"
 
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
        sSQL = "SELECT [WS_ID],[CODE_ID],[WORK ORDER],[ATC PART]," & _
                       "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
                      "[TOTAL TIME],[QUANTITY],[REJECTS] " & _
               "FROM [WORK SHEET] " & _
               "WHERE [DATE_ID]      =#" & DATE_ID & "# AND " & _
                     "[OP_ID]     =" & OP_ID & " AND " & _
                     "[MACHINE_ID]= " & MACHINE_ID & " " & _
               "ORDER BY [WS_ID] DESC"
Case 1

        sSQL = "SELECT [WS_ID],[CODE_ID],[WORK ORDER],[ATC PART]," & _
                       "format([START TIME],'h:mm AM/PM'),format([STOP TIME],'h:mm AM/PM')," & _
                      "[TOTAL TIME],[QUANTITY],[REJECTS] " & _
               "FROM [WORK SHEET] " & _
               "WHERE [OP_ID]=" & OP_ID & " AND " & _
                     "[MACHINE_ID]= " & MACHINE_ID & " " & _
               "ORDER BY [WS_ID] DESC"
End Select


sSQLF = "    ||^Code|<Product                |<Atc Part            |^Start            |^Stop     "
sSQLF = sSQLF & "    |Time  |Produced|>Defects|>Restock"

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

If (Format(DTPicker1.value, "hh:mm") > Format(DTPicker2.value, "hh:mm")) Then
    txtTotalTime.Text = DateDiff("n", Format(DTPicker1.value, "hh:mm"), Format(DTPicker2.value, "hh:mm")) + 1440
Else
    txtTotalTime.Text = DateDiff("n", Format(DTPicker1.value, "hh:mm"), Format(DTPicker2.value, "hh:mm"))
End If
 
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
sSQL = "SELECT SUM([QTY]) AS [SUM REJECTS] " & _
       "FROM [DEFECTS] " & _
       "WHERE [WS_ID]=" & WS_ID & " GROUP BY [WS_ID]"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    txtDefects.Text = FR_Table.Fields("[SUM REJECTS]")
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

Private Sub CommandAdd_Click()

If TRAY_ID <> 0 And POWER_ID <> 0 Then
    VALID_PART_ID = 1
Else
    VALID_PART_ID = 0
End If

Select Case VALID_PART_ID
Case 1
    
    Work_Codes TRAY_ID + 750, LabelTRAY_ID.Caption, 0, frmWorkSheet1
    
    cmdRefreshDisplay_Click
    MSFlexGrid1_Click
    
    txtATCPart.Text = TextATCPart.Text
    ATC_PART_ID = TextATCPart.Text
    txtWorkOrder.Text = TextWorkOrder.Text
    Data4.UpdateRecord
    
    cmdRefreshDisplay_Click
    MSFlexGrid1_Click
    
End Select


End Sub

Private Sub CommandSet_Click()

ValidPartNew (TextATCPart.Text)

LabelSeriesCase.Caption = Mid(ATC_PART_ID, 1, 4)
LabelDV_ID.Caption = DV_ID
LabelTS.Caption = Mid(ATC_PART_ID, 9, 2)

Tray_Power_Lookup ATC_PART_ID

lblTRAY_ID.Caption = TRAY_ID
LabelPOWER_ID.Caption = POWER_ID

Dim ValidChar As String
Dim SearchChar As String
Dim MYPOS As Integer
Dim found As Integer

Select Case Mid(ATC_PART_ID, 4, 1)
Case "E"
            
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
                        TRAY_ID = 1
                        found = 1
                End Select
        End Select
    
        If (found = 0) Then
                Select Case Mid(ATC_PART_ID, 9, 1)
                Case "Q", "R", "O"
                        TRAY_ID = 6
                Case "C", "I", "P", "S", "T", "W", "Y"
                        TRAY_ID = 5
                Case "A", "G", "J", "U"
                        TRAY_ID = 7
                Case "D", "E", "K", "M"
                        Select Case Mid(ATC_PART_ID, 11, 1)
                        Case "X"
                            TRAY_ID = 10
                        Case Else
                            TRAY_ID = 11
                        End Select
                End Select
        End If
        Option1.value = True
        FixturePage
 
Case "B"
        TRAY_ID = 8
        Option3.value = True
        FixturePage
Case "C"
        Option2.value = True
        Select Case Mid(ATC_PART_ID, 9, 1)
        Case "Q", "R", "O"
                    TRAY_ID = 3
        Case "C", "I", "P", "S", "T", "W", "Y"
                    TRAY_ID = 2
        Case "A", "G", "J", "U"
                    TRAY_ID = 4
        Case "D", "E", "K", "M"
                Select Case Mid(ATC_PART_ID, 11, 1)
                Case "X"
                    TRAY_ID = 12
                Case Else
                    TRAY_ID = 4
                End Select
        End Select
        Option2.value = True
        FixturePage
Case "H"
        Option4.value = True
        FixturePage
        TRAY_ID = 13
Case Else

End Select

'Text1.Text = UCase(Mid(txtATCPart.Text, 5, 4))

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

Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String

sSQL = "SELECT [TRAY_ID],[ROWS] &  'x' & [COLS] AS [SQL NAME],[TITLE],[ATC DWG]  " & _
       "FROM [TBL TRAY CONFIG] WHERE [TRAY_ID] =" & TRAY_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)


If (FR_Table.RecordCount <> 0) Then
        sSQL = " TRAY_ID " & TRAY_ID & "      [" & FR_Table.Fields("[TRAY_ID]") & "] " & FR_Table.Fields("[SQL NAME]") & " " & FR_Table.Fields("[TITLE]") & " " & FR_Table.Fields("[ATC DWG]")

        lblCode_ID.Caption = TRAY_ID + 750
Else
        sSQL = ""
        lblCode_ID.Caption = ""
End If

FR_Table.Close
FR_Database.Close

LabelTRAY_ID.Caption = sSQL

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
Data7.DatabaseName = ATC_LASER_BD
 
MSFlexGrid1.Top = 0
MSFlexGrid1.Width = 11600
MSFlexGrid1.Height = 4000
MSFlexGrid1.ForeColorSel = vbBlack

MSFlexGrid7.Width = 11600
MSFlexGrid7.Height = 2800
MSFlexGrid7.ForeColorSel = vbBlack

MSFlexGrid2.Width = 2200
MSFlexGrid2.Height = 2400

MSFlexGrid3.Width = 3400
MSFlexGrid3.Height = 1200

lblNote.Caption = "Note : ATC Part Format chr 10 Non Mag IN    (N,1,3,5,7,9,C,H,B)"

Dim sSQL As String
Dim sSQLF As String

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    MACHINE_TYPE = FR_Table.Fields("[TYPE NUMBER]")
    MACHINE_DESCRIPTION = FR_Table.Fields("[DESCRIPTION]")
    DEPT_ID = "LS"
    lblMachine.Caption = FR_Table.Fields("[MACHINE]") & " " & FR_Table.Fields("[DESCRIPTION]")
End If

sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID]=" & OP_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    txtOperator.Caption = FR_Table.Fields("[FIRST]") & " " & FR_Table.Fields("[LAST]")
    txtShift.Caption = FR_Table.Fields("[SHIFT_ID]")
End If
FR_Table.Close
FR_Database.Close

'=============================================================================
'   LOAD DEFECT LIST
'=============================================================================)

sSQL = "SELECT [DEFECT_ID],[DESCRIPTION]  " & _
        "FROM [DEFECT LIST] WHERE [LS]='L'"
 
sSQLF = "    ||<Defect Description"
 
Data2.RecordSource = sSQL
Data2.Refresh

MSFlexGrid2.FormatString = sSQLF
 
'=============================================================================
'   LOAD DEPT CODES
'=============================================================================
   
FixturePage
                
cmdRefresh2_Click
        
DTPicker3.value = Date$
DATE_ID = Date$
    
TextWorkOrder.Text = "Work Order"

TextATCPart.Text = "ATC Part"
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

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then

End If

End Sub

Private Sub MSFlexGrid1_Click()

MSFlexGrid1.Col = 1
WS_ID = Val(MSFlexGrid1.Text)

'If (WS_ID <> 0) Then
'    fraWS.Enabled = True
'Else
'    fraWS.Enabled = False
'End If

MSFlexGrid1.Col = 2
CODE_ID = Val(MSFlexGrid1.Text)

TRAY_ID = CODE_ID - 750

Set FR_Database = OpenDatabase(ATC_LASER_BD)
Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT * FROM [TBL TRAY CONFIG] " & _
       "WHERE [TRAY_ID] = " & TRAY_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    CASE_ID = FR_Table.Fields("[CASE]")
    TRAY_ID = FR_Table.Fields("[TRAY_ID]")
Else

End If
FR_Database.Close

fraWS.Caption = "CODE_ID : " & CODE_ID

MSFlexGrid1.Col = 4
ATC_PART_ID = MSFlexGrid1.Text

sSQL = "SELECT * FROM [WORK SHEET] WHERE [WS_ID]=" & WS_ID
Data4.RecordSource = sSQL
Data4.Refresh
        
sSQL = "SELECT [TBL_ID],[SERIES],[CASE],[VALUE],[COATING],[ATC PART]," & _
               "format([Frequency],'0.00')," & _
               "format([Markspeed],'0.000')," & _
               "format([PulseWidth],'0.00')" & _
       "FROM [TBL Power] " & _
       "WHERE [ACTIVE] = Yes AND " & _
             "[CASE]='" & CASE_ID & "' AND " & _
             "[TRAY_ID]=" & TRAY_ID
                                   
Select Case TRAY_ID
Case 8
        sSQLF = "   |^PWR_ID|<Series        ||^DV Range       |<Coating          |<ATC Part                |>Freq     |>MKS       |>PW      "
Case Else
        sSQLF = "   |^PWR_ID|<Series   |^Case|^DV Range       |<Coating          |<ATC Part                |>Freq     |>MKS       |>PW      "
End Select

Data7.RecordSource = sSQL
Data7.Refresh
MSFlexGrid7.FormatString = sSQLF
        
cmdRefresh3_Click
        
cmdRefresh2_Click
MSFlexGrid6_Click

'MSFlexGrid7_Click   'POWER_ID

MSFlexGrid6_Click   'SIZE_LOC_ID

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

Private Sub MSFlexGrid6_Click()

MSFlexGrid6.Col = 1
SIZE_LOC_ID = Val(MSFlexGrid6.Text)
 
MSFlexGrid6.Col = 0
MSFlexGrid6.ColSel = MSFlexGrid6.Cols - 1 '10

End Sub

Private Sub MSFlexGrid7_Click()

MSFlexGrid7.Col = 1
POWER_ID = Val(MSFlexGrid7.Text)
 
MSFlexGrid7.Col = 0
MSFlexGrid7.ColSel = MSFlexGrid7.Cols - 1 '10

End Sub

Private Sub Option1_Click()
FixturePage
End Sub

Private Sub Option2_Click()
FixturePage
End Sub

Private Sub Option3_Click()
FixturePage
End Sub

Private Sub Option4_Click()
FixturePage
End Sub

Private Sub TextATCPart_GotFocus()
TextATCPart.SelStart = 0
TextATCPart.SelLength = Len(TextATCPart)
End Sub

Private Sub TextATCPart_LostFocus()

TextATCPart.Text = Trim(TextATCPart)
TextATCPart.Text = Mid(UCase(TextATCPart.Text), 1, 12)

'
'   SCHEDULE LOOKUP BY WORK ORDER
'
Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)
Dim sSQL As String
sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='" & TextATCPart.Text & "'"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    If (FR_Table.Fields("[ATC PART]") <> vbNull) Then
        TextATCPart.Text = FR_Table.Fields("[ATC PART]")
        ATC_PART_ID = FR_Table.Fields("[ATC PART]")
    End If
End If
FR_Table.Close
FR_Database.Close

CommandSet_Click

End Sub

Private Sub TextWorkOrder_GotFocus()
TextWorkOrder.SelStart = 0
TextWorkOrder.SelLength = Len(TextWorkOrder)
End Sub

Private Sub TextWorkOrder_LostFocus()

TextWorkOrder.Text = Trim(TextWorkOrder.Text)
TextWorkOrder.Text = Mid(UCase(TextWorkOrder.Text), 1, 12)
'
'   SCHEDULE LOOKUP BY WORK ORDER
'
Set FR_Database = OpenDatabase(DB_MASTER_SCHEDULE)

Dim sSQL As String
sSQL = "SELECT * FROM [WO SCHED] WHERE [WORK ORDER]='" & TextWorkOrder.Text & "'"
              
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
If (FR_Table.RecordCount <> 0) Then
    If IsNull(FR_Table.Fields("[ATC PART]")) = False Then
        TextATCPart.Text = FR_Table.Fields("[ATC PART]")
        ATC_PART_ID = FR_Table.Fields("[ATC PART]")
    End If

End If
FR_Table.Close
FR_Database.Close
 
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
        ValidChar = "N123579CHB"
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

Private Sub txtLot_GotFocus()
txtLot.SelStart = 0
txtLot.SelLength = Len(txtLot)
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
            txtSQ.Text = FR_Table.Fields("[START QTY]")
        Else
            txtOrderQty.Text = 0
            txtSQ.Text = 0
        End If
    End If
    FR_Table.Close
    FR_Database.Close

End If
 
'Data4.UpdateRecord

cmdRefreshDisplay_Click

End Sub


Public Sub Work_Codes(iCode As Long, sDescription As String, iTotalTime As Integer, frmWorkSheet As Form)

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT * FROM [WORK SHEET] " & _
       "WHERE [DATE_ID]    =#" & DATE_ID & "# AND " & _
             "[OP_ID]      =" & OP_ID & " AND " & _
             "[MACHINE_ID] = " & MACHINE_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
FR_Table.AddNew
WS_ID = FR_Table.Fields("WS_ID")

FR_Table.Fields("[OP_ID]") = OP_ID
FR_Table.Fields("[DATE_ID]") = DATE_ID
FR_Table.Fields("[MACHINE_ID]") = MACHINE_ID
FR_Table.Fields("[CODE_ID]") = iCode
If (Len(sDescription) > 14) Then
    sDescription = Mid(sDescription, 1, 14)
End If

FR_Table.Fields("[WORK ORDER]") = Mid(sDescription, 1, 12)
FR_Table.Fields("[ATC PART]") = "ATCPART"
FR_Table.Fields("[LOT NUM]") = "Lot Number"
FR_Table.Fields("[QUANTITY]") = 0
FR_Table.Fields("[RESTOCK]") = 0
FR_Table.Fields("[REJECTS]") = 0
FR_Table.Fields("[START TIME]") = Format(Date, "MM/dd/yyyy ") & Format(Time, "hh:mm am/pm")
FR_Table.Fields("[TOTAL TIME]") = 0

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

Public Sub FixturePage()

Dim IPAGE As Integer

If (Option1.value = True) Then
    IPAGE = 1
End If
If (Option2.value = True) Then
    IPAGE = 2
End If
If (Option3.value = True) Then
    IPAGE = 3
End If
If (Option4.value = True) Then
    IPAGE = 4
End If

Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String

sSQL = "SELECT [TRAY_ID],[ROWS] &  'x' & [COLS] AS [SQL NAME],[TITLE],[ATC DWG]  " & _
       "FROM [TBL TRAY CONFIG] " & _
       "WHERE [ACTIVE]=1 AND [PAGE] =" & IPAGE & _
       " ORDER BY [CASE]"

Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim i As Integer

For i = 0 To 5
        lblCode(i).Visible = False
        cmdCode(i).Visible = False
Next i

i = 0
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        lblCode(i).Caption = FR_Table.Fields("[TRAY_ID]") + 750
        sSQL = "[" & FR_Table.Fields("[TRAY_ID]") & "] " & FR_Table.Fields("[SQL NAME]") & " " & FR_Table.Fields("[TITLE]") & " " & FR_Table.Fields("[ATC DWG]")
             
        cmdCode(i).Caption = sSQL
        lblCode(i).Visible = True
        cmdCode(i).Visible = True
        i = i + 1
        FR_Table.MoveNext
    Loop
End If

FR_Table.Close
FR_Database.Close

End Sub
