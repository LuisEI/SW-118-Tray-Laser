VERSION 5.00
Begin VB.Form frmMotion 
   Caption         =   "NI Motion Functions"
   ClientHeight    =   10050
   ClientLeft      =   225
   ClientTop       =   795
   ClientWidth     =   16620
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   16620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandGotoCalibration 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Goto Calibration"
      Height          =   300
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   3000
      Width           =   1665
   End
   Begin VB.CommandButton CommandSetCalibration 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Set Calibration"
      Height          =   300
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3360
      Width           =   1665
   End
   Begin VB.TextBox TextCalibration 
      BackColor       =   &H00FFC0FF&
      DataField       =   "HOME"
      DataSource      =   "Data3"
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
      Left            =   12240
      TabIndex        =   75
      Text            =   "500"
      ToolTipText     =   "HOME"
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton CommandUpdateHeight 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Set Global Height"
      Height          =   300
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3720
      Width           =   1665
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Load Parameters"
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
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   7200
      Width           =   2000
   End
   Begin VB.TextBox TextAxis3 
      BackColor       =   &H00FFC0FF&
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
      Left            =   6480
      TabIndex        =   70
      Text            =   "2000"
      Top             =   8640
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Move"
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
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   8640
      Width           =   975
   End
   Begin VB.TextBox TextAxis2 
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
      Left            =   6480
      TabIndex        =   68
      Text            =   "500"
      Top             =   8160
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Move"
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
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox TextAxis1 
      DataField       =   "HOME"
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
      Left            =   6480
      TabIndex        =   66
      Text            =   "10000"
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Move"
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
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Axis 3  Reverse Limit"
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
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "FindReverseLimit"
      Top             =   8640
      Width           =   2000
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Axis 2  Reverse Limit"
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
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "FindReverseLimit"
      Top             =   8160
      Width           =   2000
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Axis 1  Reverse Limit"
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
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "FindReverseLimit"
      Top             =   7680
      Width           =   2000
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   8880
      TabIndex        =   46
      Top             =   6240
      Width           =   7455
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Read Positions"
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
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "flex_read_pos_rtn"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFFF&
         Cancel          =   -1  'True
         Caption         =   "Read Axis Status"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "flex_read_axis_status"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblStatusH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
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
         Index           =   1
         Left            =   4560
         TabIndex        =   61
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblStatusL 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   60
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblPosition 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Limit"
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
         TabIndex        =   58
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Home"
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
         Left            =   4560
         TabIndex        =   57
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblStatusL 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
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
         Index           =   2
         Left            =   5880
         TabIndex        =   56
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblStatusH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
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
         Index           =   2
         Left            =   4560
         TabIndex        =   55
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblStatusL 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
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
         Index           =   3
         Left            =   5880
         TabIndex        =   54
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblStatusH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
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
         Index           =   3
         Left            =   4560
         TabIndex        =   53
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Position"
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
         Left            =   2400
         TabIndex        =   52
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblPosition 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   51
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblPosition 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
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
         Index           =   3
         Left            =   2400
         TabIndex        =   50
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Axis 3 : Laser Height"
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
         Left            =   360
         TabIndex        =   49
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Axis 2 : Camera"
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
         Left            =   360
         TabIndex        =   48
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Axis 1 : CarrierTray"
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
         Left            =   360
         TabIndex        =   47
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.TextBox txtTargetPosTest 
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
      TabIndex        =   17
      Text            =   "-5000"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveToTarget 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MoveToTarget"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Data Data3 
      Caption         =   "Data3 [TBL Axis]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Axis"
      Top             =   1170
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 [TBL Axis]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Axis"
      Top             =   825
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.TextBox txtVelocity 
      DataField       =   "Velocity"
      DataSource      =   "Data3"
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
      Left            =   3120
      TabIndex        =   9
      Text            =   "30.0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtAccel 
      DataField       =   "Accel"
      DataSource      =   "Data3"
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
      Left            =   5160
      TabIndex        =   10
      Text            =   "10.0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtDecel 
      DataField       =   "Decel"
      DataSource      =   "Data3"
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
      Left            =   7440
      TabIndex        =   11
      Text            =   "10.0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtVelocity 
      DataField       =   "Velocity"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   6
      Text            =   "30.0"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtAccel 
      DataField       =   "Accel"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   7
      Text            =   "10.0"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtDecel 
      DataField       =   "Decel"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   8
      Text            =   "10.0"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 [TBL Axis]"
      Connect         =   "Access"
      DatabaseName    =   "C:\ATC\104 LASER MATRIX.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TBL Axis"
      Top             =   480
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enable Home"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "flex_reset_pos"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Disable Home"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "flex_reset_pos"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "flex_stop_motion"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      Left            =   3240
      TabIndex        =   20
      Text            =   "0"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "flex_reset_pos"
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "[3]  Reverse Limit"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "flex_find_reference"
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Forward Limit"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "[1]  Initialize Controller"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Home"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Test"
      Height          =   250
      Left            =   10560
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Start to Target"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "[2]  Load Parameters"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtDecel 
      DataField       =   "Decel"
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
      Left            =   7440
      TabIndex        =   5
      Text            =   "10.0"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtTargetPos 
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
      Left            =   3240
      TabIndex        =   15
      Text            =   "5000"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtAccel 
      DataField       =   "Accel"
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
      Left            =   5160
      TabIndex        =   4
      Text            =   "10.0"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtVelocity 
      DataField       =   "Velocity"
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
      Left            =   3120
      TabIndex        =   3
      Text            =   "30.0"
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox cmbAxis 
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
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtBoardID 
      Alignment       =   2  'Center
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
      Left            =   3120
      TabIndex        =   26
      Text            =   "1"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Axis 1Calib Position"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7200
      TabIndex        =   81
      Top             =   7800
      Width           =   1365
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Laser Height Calibration"
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
      Height          =   240
      Left            =   11400
      TabIndex        =   80
      Top             =   2160
      Width           =   2520
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Move Axis 3"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7200
      TabIndex        =   79
      Top             =   8760
      Width           =   870
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "From Move Axis 3"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   12960
      TabIndex        =   78
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calibration Stick 10.235 inches"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   10440
      TabIndex        =   74
      Top             =   2520
      Width           =   2445
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4110
      Left            =   14400
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label LabelMotion 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5400
      TabIndex        =   72
      Top             =   6840
      Width           =   1605
   End
   Begin VB.Label Label14 
      Caption         =   "flex_load_target_pos"
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
      Left            =   840
      TabIndex        =   45
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "flex_find_home"
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
      Left            =   840
      TabIndex        =   44
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "flex_find_reference"
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
      Left            =   840
      TabIndex        =   43
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Axis 3 : Laser Height"
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
      Left            =   840
      TabIndex        =   42
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Axis 2 : Camera"
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
      Left            =   840
      TabIndex        =   41
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Axis 1 : CarrierTray"
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
      Left            =   840
      TabIndex        =   40
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblStatus4 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13440
      TabIndex        =   39
      Top             =   5400
      Width           =   1605
   End
   Begin VB.Label lblStatus3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6360
      TabIndex        =   38
      Top             =   5280
      Width           =   525
   End
   Begin VB.Label lblStatus2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   37
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Label lblStatus1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   36
      Top             =   4440
      Width           =   1485
   End
   Begin VB.Label lblFound 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Found"
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
      Left            =   13440
      TabIndex        =   35
      Top             =   4920
      Width           =   1605
   End
   Begin VB.Label lblFind 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Find"
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
      Left            =   13440
      TabIndex        =   34
      Top             =   4440
      Width           =   1605
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   32
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   31
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Target Position"
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
      Index           =   3
      Left            =   3240
      TabIndex        =   30
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Acceleration (RPS/s)"
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
      Index           =   4
      Left            =   5160
      TabIndex        =   29
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   28
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "BoardID"
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
      Index           =   0
      Left            =   3120
      TabIndex        =   27
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////
' Simple One-Axis Move :A Visual Basic application
'
' Requirements: FlexMotion Software Version 5 or later.
'      Use module files:
'
'      <driveletter>:\Program Files\National Instruments\Motion\FlexMotion
'          \include\*.bas
'      <driveletter>:\Program Files\National Instruments\Motion\FlexMotion
'          \example\NIMCExample.bas
'
' Description:
'      This Visual Basic example demonstrates a one axis move using motion.
'
' Note:
'   If you notice the motor is moving slowly or not moving at all, adjust
'   the Velocity, Acceleration, and Deceleration values. The default
'   values may not be suitable for your system.'///////////////////////////////////////////////////////////////////////

Option Explicit 'All Variables have to be declared
'/////////////////////////////////////////////////////////////////
' Global Variables
Dim boardID As Integer      'Board ID
Dim axis As Integer         'Axis Number
Dim targetPosition As Long  'Target Position
Dim csr As Integer          'Communication Status Register
Dim moveComplete As Integer 'Move complete status
Dim program As Integer
Dim targetPos As Long
Dim velocity As Double
Dim accel As Double
Dim decel As Double

'Global Modal variables
Dim errorCode As Long       'Modal ErrorCode
Dim commandID As Integer    'Command ID for modal error handling
Dim resourceID As Integer   'Resource ID for modal error handling
Dim error As Long


Private Sub cmdMoveToTarget_Click()

axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)

lblStatus1.Caption = " "

MoveToTarget axis, CLng(txtTargetPosTest.Text)

lblStatus1.Caption = "Complete"

End Sub

Private Sub Command1_Click()

Data1.UpdateRecord
Data2.UpdateRecord
Data3.UpdateRecord

Load_Parameters
     
MsgBox "NI Motion Parameters Complete", vbInformation, "ATC Tray Laser System"
     
End Sub

Private Sub Command10_Click()

On Error GoTo Errorhandler
      
    Dim position As Long
    
    Dim found As Integer
    Dim finding As Integer
    Dim axisStatus As Integer
      
    Screen.MousePointer = vbHourglass
      
    boardID = txtBoardID.Text
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
                                
    Dim inputVector
    inputVector = 32
                
    error = flex_find_reference(boardID, axis, 0, NIMC_FIND_REVERSE_LIMIT_REFERENCE)
        
    CheckError (error)


    Do
        error = flex_read_pos_rtn(boardID, axis, position)
        CheckError (error)
        
        error = flex_check_reference(boardID, axis, 0, found, finding)
        CheckError (error)
        
        lblFind.Caption = "Find " & finding
        lblFind.Refresh
        
        lblFound.Caption = "Found " & found
        lblFound.Refresh
        
        error = flex_read_axis_status_rtn(boardID, axis, axisStatus)
        CheckError (error)
        
        
        'Check the modal errors
       ' flex_read_csr_rtn boardID, csr
       ' If (csr And NIMC_MODAL_ERROR_MSG) Then
            'Stop the Motion
        '    flex_stop_motion boardID, axis, NIMC_DECEL_STOP, 0
         '   flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
         '   CheckError (errorCode)
        'End If
        DoEvents
    Loop Until (finding = 0)


    error = flex_reset_pos(boardID, axis, 0, 0, inputVector)
    CheckError (error)


    MsgBox "flex_find_reference Reverse Limit " & position, vbInformation, "ATC Tray Laser System"

        lblFind.Caption = "Find " & finding
        lblFind.Refresh
        
        lblFound.Caption = "Found " & found
        lblFound.Refresh

Screen.MousePointer = vbDefault

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
             '   //Read the current position of axis
        'err = flex_read_pos_rtn(boardID, axis, &position);
        'CheckError;
        ' error = flex_read_pos_rtn(boardID, axis, position)
        
         'CheckError (error)
                                
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If
    
End Sub

Private Sub Command11_Click()

On Error GoTo Errorhandler
    boardID = txtBoardID.Text
    
    Dim inputVector
    inputVector = 32
    
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
            
    Dim position As Long
    position = Val(Text1.Text)
    
    lblStatus1.Caption = " "
    
    error = flex_reset_pos(BOARD_ID, axis, position, position, inputVector)
    CheckError (error)

    'Check the modal errors
    If csr And NIMC_MODAL_ERROR_MSG Then
        error = flex_stop_motion(BOARD_ID, axis, NIMC_DECEL_STOP, 0) 'Stop the Motion
        flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
        CheckError (errorCode)
    End If

    lblStatus1.Caption = "Complete"
    'MsgBox "National Instruments flex_reset_pos Complete", vbInformation, "ATC Tray Laser System"
        
    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

Private Sub Command12_Click()

On Error GoTo Errorhandler
    
    boardID = txtBoardID.Text
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
  
 'flex_enable_home_inputs Lib "FlexMotion32.dll" (ByVal boardID%, ByVal homemap%) As Long
        
    Dim homemap%
        
    homemap% = 0
                
    homemap% = BinaryToDecimal("0")
    
    error = flex_enable_home_inputs(boardID, homemap%)
       
    CheckError (error)
   
    MsgBox "Complete", vbInformation, "ATC Tray Laser System"

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If
End Sub

Private Sub Command13_Click()
On Error GoTo Errorhandler
    
    boardID = txtBoardID.Text
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
  
 'flex_enable_home_inputs Lib "FlexMotion32.dll" (ByVal boardID%, ByVal homemap%) As Long
        
    Dim homemap%
        
    homemap% = 14
                
    homemap% = BinaryToDecimal("1110")
    
    error = flex_enable_home_inputs(boardID, homemap%)
       
    CheckError (error)
   
    MsgBox "Complete", vbInformation, "ATC Tray Laser System"

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If
    
End Sub

Private Sub Command14_Click()

FindReverseLimit 1

End Sub

Private Sub Command15_Click()

FindReverseLimit 2

End Sub

Private Sub Command16_Click()

FindReverseLimit 3

End Sub

Private Sub Command17_Click()
 
LabelMotion.Caption = "Start"
Screen.MousePointer = vbHourglass

MoveToTarget 1, CLng(TextAxis1.Text)

Screen.MousePointer = vbDefault
LabelMotion.Caption = "Complete"

End Sub

Private Sub Command18_Click()

LabelMotion.Caption = "Start"
Screen.MousePointer = vbHourglass

MoveToTarget 2, CLng(TextAxis2.Text)

Screen.MousePointer = vbDefault
LabelMotion.Caption = "Complete"

End Sub

Private Sub Command19_Click()

LabelMotion.Caption = "Start"
Screen.MousePointer = vbHourglass

MoveToTarget 3, CLng(TextAxis3.Text)

Screen.MousePointer = vbDefault
LabelMotion.Caption = "Complete"

End Sub

Private Sub Command2_Click()

  On Error GoTo Errorhandler
    'Set the boardID,axis and targetPosition
    boardID = CInt(txtBoardID.Text)
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
    targetPosition = CLng(txtTargetPos.Text)
    
    'Load the operation mode - absolute position
    error = flex_set_op_mode(boardID, axis, NIMC_ABSOLUTE_POSITION)
 
    'Load a target position of 20000 counts or steps
    error = flex_load_target_pos(boardID, axis, targetPosition, &HFF)
    
    'Start the motion
    error = flex_start(boardID, axis, 0)
    
    lblStatus1.Caption = " "
    
    Do
        'Check the move complete status
        error = flex_check_move_complete_status(boardID, axis, 0, moveComplete)
        
        'Check the modal errors
        flex_read_csr_rtn boardID, csr
        If (csr And NIMC_MODAL_ERROR_MSG) Then
            'Stop the Motion
            flex_stop_motion boardID, axis, NIMC_DECEL_STOP, 0
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            CheckError (errorCode)
        End If
        
        DoEvents
    Loop Until moveComplete
    
    lblStatus1.Caption = "Complete"
  
    
    Exit Sub    'Exit the Sub
            
'////////////////////////////////////////////////////////////////////////
' Error Handling
'
Errorhandler:
    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

 
Private Sub Command20_Click()

Initialize_Controller

DisableHome

Load_Parameters

MsgBox "NI Motion Parameters Complete", vbInformation, "ATC Tray Laser System"
End Sub

Private Sub Command3_Click()

On Error GoTo Errorhandler

    lblStatus3.Caption = ""
    'Set the boardID,axis and targetPosition
    boardID = CInt(txtBoardID.Text)
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
 
            'Stop the Motion
    flex_stop_motion boardID, axis, NIMC_DECEL_STOP, 0
            
    flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
                
    'MsgBox "National Instruments Stop Motion Complete", vbInformation, "ATC Tray Laser System"
     
    lblStatus3.Caption = "OK"
     
    Exit Sub    'Exit the Sub
            
'////////////////////////////////////////////////////////////////////////
' Error Handling
'
Errorhandler:
    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

Private Sub Command4_Click()

On Error GoTo Errorhandler
    
    boardID = txtBoardID.Text
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
     
    ' flex_set_limit_input_polarity Lib "FlexMotion32.dll" (ByVal boardID%, ByVal forwardPolarityMap%, ByVal reversePolarityMap%) As Long
    Dim forwardPolarityMap As Long
    Dim reversePolarityMap As Long
   
    forwardPolarityMap = 1 + 2 + 4 + 8
    reversePolarityMap = 1 + 2 + 4 + 8
       
    forwardPolarityMap = 0
    reversePolarityMap = 0
    
    error = flex_set_limit_input_polarity(boardID, forwardPolarityMap, reversePolarityMap)
    
    CheckError (error)

    MsgBox "Test Flex ", vbInformation, "ATC Tray Laser System"

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If
    
End Sub

Private Sub Command5_Click()

On Error GoTo Errorhandler
    
    boardID = txtBoardID.Text
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
        
    lblStatus3.Caption = ""
        
    'FIND REFERENCE
    'flex_find_home Lib "FlexMotion32.dll" (ByVal boardID%, ByVal axis%, ByVal directionMap%) As Long
        
    error = flex_find_home(boardID, axis, 0)
       
    CheckError (error)

    lblStatus3.Caption = "Home"

    'MsgBox "Find Home Complete", vbInformation, "ATC Tray Laser System"

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

Private Sub Command6_Click()

Initialize_Controller

MsgBox "National Instruments Motion Initialed", vbInformation, "ATC Tray Laser System"

End Sub

Private Sub Command7_Click()
    
On Error GoTo Errorhandler
      
    Dim position As Long
    
    Dim found As Integer
    Dim finding As Integer
    Dim axisStatus As Integer
      
    boardID = txtBoardID.Text
    axis = cmbAxis(1).ItemData(cmbAxis(1).ListIndex)
                
    error = flex_find_reference(boardID, axis, 0, NIMC_FIND_FORWARD_LIMIT_REFERENCE)
        
    CheckError (error)

    Do
        error = flex_read_pos_rtn(boardID, axis, position)
        CheckError (error)
        
        error = flex_check_reference(boardID, axis, 0, found, finding)
        CheckError (error)
        
        lblFind.Caption = "Find " & finding
        lblFind.Refresh
        
        lblFound.Caption = "Found " & found
        lblFound.Refresh
        
        error = flex_read_axis_status_rtn(boardID, axis, axisStatus)
        CheckError (error)
                
        'Check the modal errors
       ' flex_read_csr_rtn boardID, csr
       ' If (csr And NIMC_MODAL_ERROR_MSG) Then
            'Stop the Motion
        '    flex_stop_motion boardID, axis, NIMC_DECEL_STOP, 0
         '   flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
         '   CheckError (errorCode)
        'End If
        DoEvents
    Loop Until (finding = 0)

    MsgBox "flex_find_reference Reverse Limit " & position, vbInformation, "ATC Tray Laser System"

    lblFind.Caption = "Find " & finding
    lblFind.Refresh
    
    lblFound.Caption = "Found " & found
    lblFound.Refresh

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
             '   //Read the current position of axis
        'err = flex_read_pos_rtn(boardID, axis, &position);
        'CheckError;
        ' error = flex_read_pos_rtn(boardID, axis, position)
        
         'CheckError (error)
                                
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

    
    
End Sub

Private Sub Command8_Click()
On Error GoTo Errorhandler

    'Set the boardID,axis and targetPosition
    boardID = CInt(txtBoardID.Text)
         
    Dim axisStatus%
    Dim sBuff As String
                                        
    For axis = 1 To 3
            
            flex_read_axis_status_rtn boardID, axis, axisStatus%
                                                   
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
                        
            Select Case 3
            Case 0
                    lblStatusH(axis).Caption = axisStatus%
            Case 1
                    lblStatusH(axis).Caption = Hex(axisStatus%)
            Case 2
                    lblStatusH(axis).Caption = Oct(axisStatus%)
            Case 3
                    sBuff = LongToBinary(axisStatus%)
                    'lblStatus.Caption = Mid(sBuff, Len(sBuff) - 16, Len(sBuff))
                    
                    lblStatusH(axis).Caption = Mid(sBuff, Len(sBuff) - 5, 1)
                    lblStatusL(axis).Caption = Mid(sBuff, Len(sBuff) - 4, 1)
                    
                    
                    DoEvents
            End Select
    Next axis
    
     
    Exit Sub    'Exit the Sub
            
'////////////////////////////////////////////////////////////////////////
' Error Handling
'
Errorhandler:
    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If
End Sub
 

Private Sub Command9_Click()

On Error GoTo Errorhandler

    Dim position As Long

    boardID = txtBoardID.Text
   
   For axis = 1 To 3
    error = flex_read_pos_rtn(boardID, axis, position)
        CheckError (error)
    
        'Check the modal errors
        If csr And NIMC_MODAL_ERROR_MSG Then
            error = flex_stop_motion(boardID, axis, NIMC_DECEL_STOP, 0) 'Stop the Motion
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            CheckError (errorCode)
        End If
        
        lblPosition(axis).Caption = position
        lblPosition(axis).Refresh
        DoEvents
        
   Next axis
                        
    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn boardID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn boardID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn boardID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

Private Sub CommandGotoCalibration_Click()

LabelMotion.Caption = "Start"
Screen.MousePointer = vbHourglass

MoveToTarget 1, CLng(TextAxis1.Text)

MoveToTarget 3, CLng(TextCalibration.Text)

Screen.MousePointer = vbDefault
LabelMotion.Caption = "Complete"

End Sub

Private Sub CommandSetCalibration_Click()

TextCalibration.Text = TextAxis3.Text

End Sub

Private Sub CommandUpdateHeight_Click()

Dim CAL_HT_POSITION As Long

CAL_HT_POSITION = TextCalibration.Text

Dim RATIO_STEPS_PER_INCH As Double

RATIO_STEPS_PER_INCH = 2000

Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String
  
sSQL = "SELECT * FROM [TBL Power] WHERE  [ACTIVE] = Yes  ORDER BY [ORDER]"

Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Dim COUNT As Integer
 
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
            FR_Table.Edit
            FR_Table.Fields("[Z HEIGHT]") = CAL_HT_POSITION + (FR_Table.Fields("[HEIGHT]") * RATIO_STEPS_PER_INCH)
            FR_Table.Update
            FR_Table.MoveNext
            COUNT = COUNT + 1
    Loop
End If
FR_Table.Close
FR_Database.Close

MsgBox "Complete COUNT " & COUNT, vbInformation, "ATC Data Base System"

End Sub

'/////////////////////////////////////////////////////////////////
' Form Load - Initializations
Private Sub Form_Load()
    

Caption = "NI Motion Functions   " & ATC_DWG & "         " & ATC_VERSION


cmbAxis(1).ListIndex = 0    'Set the combo box to axis 1

Dim sSQL As String

Data1.DatabaseName = ATC_LASER_BD

sSQL = "SELECT * FROM  [TBL Axis] WHERE [AXIS_ID] = 1"
                    
Data1.RecordSource = sSQL
Data1.Refresh

Data2.DatabaseName = ATC_LASER_BD

sSQL = "SELECT * FROM  [TBL Axis] WHERE [AXIS_ID] = 2"
                    
Data2.RecordSource = sSQL
Data2.Refresh

Data3.DatabaseName = ATC_LASER_BD

sSQL = "SELECT * FROM  [TBL Axis] WHERE [AXIS_ID] = 3"
                    
Data3.RecordSource = sSQL
Data3.Refresh

lblFind.Caption = ""
lblFound.Caption = ""

TextAxis3.Text = TextCalibration.Text

End Sub


 Private Function LongToBinary(ByVal long_value As Long, _
    Optional ByVal separate_bytes As Boolean = True) As _
    String
Dim hex_string As String
Dim digit_num As Integer
Dim digit_value As Integer
Dim nibble_string As String
Dim result_string As String
Dim factor As Integer
Dim bit As Integer

    ' Convert into hex.
    hex_string = Hex$(long_value)

    ' Zero-pad to a full 8 characters.
    hex_string = Right$(String$(8, "0") & hex_string, 8)

    ' Read the hexadecimal digits
    ' one at a time from right to left.
    For digit_num = 8 To 1 Step -1
        ' Convert this hexadecimal digit into a
        ' binary nibble.
        digit_value = CLng("&H" & Mid$(hex_string, _
            digit_num, 1))

        ' Convert the value into bits.
        factor = 1
        nibble_string = ""
        For bit = 3 To 0 Step -1
            If digit_value And factor Then
                nibble_string = "1" & nibble_string
            Else
                nibble_string = "0" & nibble_string
            End If
            factor = factor * 2
        Next bit

        ' Add the nibble's string to the left of the
        ' result string.
        result_string = nibble_string & result_string
    Next digit_num

    ' Add spaces between bytes if desired.
    If separate_bytes Then
        result_string = _
            Mid$(result_string, 1, 8) & " " & _
            Mid$(result_string, 9, 8) & " " & _
            Mid$(result_string, 17, 8) & " " & _
            Mid$(result_string, 25, 8)
    End If

    ' Return the result.
    LongToBinary = result_string
End Function


Public Function BinaryToDecimal(Binary As String) As Long
Dim n As Long
Dim s As Integer

    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ _
            (s - 1)))
    Next s

    BinaryToDecimal = n
End Function

Private Sub Form_Unload(Cancel As Integer)
Select Case OP_MODE
Case 0
        frmMain.Show
Case 1
        frmOPScreen.Show
End Select
End Sub

Private Sub txtAccel_GotFocus(Index As Integer)
txtAccel(Index).SelStart = 0
txtAccel(Index).SelLength = Len(txtAccel(Index))
End Sub

Private Sub txtDecel_GotFocus(Index As Integer)
txtDecel(Index).SelStart = 0
txtDecel(Index).SelLength = Len(txtDecel(Index))
End Sub

Private Sub txtVelocity_GotFocus(Index As Integer)
txtVelocity(Index).SelStart = 0
txtVelocity(Index).SelLength = Len(txtVelocity(Index))
End Sub
