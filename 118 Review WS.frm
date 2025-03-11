VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReviewWS 
   Caption         =   "118 Review Work Sheets"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   15600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   20
      Text            =   "XXXX"
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame fraTest 
      Caption         =   " Search "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   6480
      Width           =   3495
      Begin VB.TextBox txtPart 
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
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "100B370"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdLocatePart 
         Caption         =   "ATC Part"
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
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdLocate 
         Caption         =   "Lot"
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
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtLot 
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
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "N43X7R5ASB"
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   3495
      Begin VB.CommandButton cmdWS 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Work Sheet"
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
         Height          =   360
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblOperator 
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
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2715
      End
      Begin VB.Label lblMachineNo 
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
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   555
      End
      Begin VB.Label lblDescription 
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
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
      Begin VB.CheckBox chkDetail 
         Caption         =   "[1] Detail"
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
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdRefresh 
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
         Height          =   360
         Left            =   1920
         TabIndex        =   16
         Top             =   1200
         Width           =   1350
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<< Day"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1350
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Day >>"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   1350
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "Week"
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
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton optDay 
         Caption         =   "Day"
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
         TabIndex        =   3
         Top             =   1680
         Value           =   -1  'True
         Width           =   855
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
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
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
         CustomFormat    =   "h:mm tt"
         Format          =   92995585
         CurrentDate     =   38117
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
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
         CustomFormat    =   "h:mm tt"
         Format          =   92995585
         CurrentDate     =   38117
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2 FROM [WORK SHEET],[MACHINE],[BARCODE]"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   5580
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1095
      Left            =   3960
      TabIndex        =   0
      ToolTipText     =   "FROM [WORK SHEET],[MACHINE],[BARCODE]"
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3690
   End
End
Attribute VB_Name = "frmReviewWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdLocate_Click()

Screen.MousePointer = vbHourglass

Dim LOT_ID As String
LOT_ID = txtLot.Text & "*"

Dim sSQL As String
Dim sSQLF As String
 
sSQL = "SELECT [WORK SHEET].[DATE_ID]       AS [SQL 1]," & _
                 "[MACHINE].[MACHINE_ID] AS [SQL 2]," & _
              "[WORK SHEET].[OP_ID]      AS [SQL 3]," & _
              "[BARCODE].[FIRST] &  ' ' & [BARCODE].[LAST] AS [SQL 4]," & _
              "[WORK SHEET].[LOT NUM]," & _
              "[WORK SHEET].[ATC PART]," & _
              "format([WORK SHEET].[TOTAL TIME]/60,'0.0')," & _
              "format([WORK SHEET].[QUANTITY],'###,##0')," & _
              "format([WORK SHEET].[REJECTS],'##,##0') " & _
       "FROM [WORK SHEET],[MACHINE],[BARCODE] " & _
       "WHERE  [WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
              "[WORK SHEET].[OP_ID]      = [BARCODE].[OP_ID] AND " & _
              "[WORK SHEET].[MACHINE_ID] =" & MACHINE_ID & " AND " & _
              "mid([WORK SHEET].[WORK ORDER],1,10)='" & LOT_ID & "'"


sSQL = "SELECT [WORK SHEET].[DATE_ID]       AS [SQL 1]," & _
                 "[MACHINE].[MACHINE_ID] AS [SQL 2]," & _
              "[WORK SHEET].[OP_ID]      AS [SQL 3]," & _
              "[BARCODE].[FIRST] &  ' ' & [BARCODE].[LAST] AS [SQL 4]," & _
              "[WORK SHEET].[LOT NUM]," & _
              "[WORK SHEET].[ATC PART]," & _
              "format([WORK SHEET].[TOTAL TIME]/60,'0.0')," & _
              "format([WORK SHEET].[QUANTITY],'###,##0')," & _
              "format([WORK SHEET].[REJECTS],'##,##0') " & _
       "FROM [WORK SHEET],[MACHINE],[BARCODE] " & _
       "WHERE  [WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
              "[WORK SHEET].[OP_ID]      = [BARCODE].[OP_ID] AND " & _
              "[WORK SHEET].[MACHINE_ID] =" & MACHINE_ID & " AND " & _
              "[WORK SHEET].[WORK ORDER] LIKE '" & LOT_ID & "'"


sSQLF = "    |^Date             ||"
sSQLF = sSQLF & "|<Operator                         "
sSQLF = sSQLF & "|^Lot W.O.                  |<ATC Part                         "
sSQLF = sSQLF & "|>Hours   |>Quantity|>Defects"
 
 
Data2.RecordSource = sSQL
Data2.Refresh
MSFlexGrid2.FormatString = sSQLF

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdLocatePart_Click()
Screen.MousePointer = vbHourglass

Dim LOT_ID As String
LOT_ID = txtPart.Text

Dim sSQL As String
Dim sSQLF As String

sSQL = "SELECT [WORK SHEET].[DATE_ID]       AS [SQL 1]," & _
                 "[MACHINE].[MACHINE_ID] AS [SQL 2]," & _
              "[WORK SHEET].[OP_ID]      AS [SQL 3]," & _
              "[BARCODE].[FIRST] &  ' ' & [BARCODE].[LAST] AS [SQL 4]," & _
              "[WORK SHEET].[LOT NUM]," & _
              "[WORK SHEET].[ATC PART]," & _
              "format([WORK SHEET].[TOTAL TIME]/60,'0.0')," & _
              "format([WORK SHEET].[QUANTITY],'###,##0')," & _
              "format([WORK SHEET].[REJECTS],'##,##0') " & _
       "FROM [WORK SHEET],[MACHINE],[BARCODE] " & _
       "WHERE  [WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
              "[WORK SHEET].[OP_ID]      = [BARCODE].[OP_ID] AND " & _
              "[WORK SHEET].[MACHINE_ID] =" & MACHINE_ID & " AND " & _
              "mid([WORK SHEET].[ATC PART],1,7)='" & LOT_ID & "'"


sSQLF = "    |^Date             ||"
sSQLF = sSQLF & "|<Operator                         "
sSQLF = sSQLF & "|^Lot W.O.                  |<ATC Part                         "
sSQLF = sSQLF & "|>Hours   |>Quantity|>Defects"
 
 
Data2.RecordSource = sSQL
Data2.Refresh
 
MSFlexGrid2.FormatString = sSQLF

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdNext_Click()

If (optDay.value = True) Then
    DTPicker1.value = DateAdd("d", 1, DTPicker1.value)
    DTPicker2.value = DateAdd("d", 1, DTPicker2.value)
Else
    DTPicker1.value = DateAdd("ww", 1, DTPicker1.value)
    DTPicker2.value = DateAdd("ww", 1, DTPicker2.value)
End If
cmdRefresh_Click

End Sub

Private Sub cmdPrevious_Click()

If (optDay.value = True) Then
    DTPicker1.value = DateAdd("d", -1, DTPicker1.value)
    DTPicker2.value = DateAdd("d", -1, DTPicker2.value)
Else
    DTPicker1.value = DateAdd("ww", -1, DTPicker1.value)
    DTPicker2.value = DateAdd("ww", -1, DTPicker2.value)
End If

cmdRefresh_Click

End Sub
 
Private Sub cmdRefresh_Click()

MSFlexGrid2.Enabled = True
 
DATE_START_ID = DTPicker1.value
DATE_END_ID = DTPicker2.value

Dim sSQL As String
Dim sSQLF As String

If (chkDetail.value = vbUnchecked) Then

            sSQL = "SELECT first([WORK SHEET].[DATE_ID])                AS [SQL 1]," & _
                             "first([MACHINE].[MACHINE_ID])          AS [SQL 2]," & _
                          "first([WORK SHEET].[OP_ID])               AS [SQL 3]," & _
          "first([MACHINE].[MACHINE]),first([MACHINE].[DESCRIPTION]) AS [SQL 4]," & _
                          "first([BARCODE].[FIRST]) &  ' ' & first([BARCODE].[LAST])," & _
                          "first([BARCODE].[SHIFT_ID])," & _
                          "count([WORK SHEET].[QUANTITY])," & _
                          "format(sum([WORK SHEET].[TOTAL TIME])/60,'0.0')," & _
                          "format(sum([WORK SHEET].[QUANTITY]),'@@@,@@@')," & _
                          "format(sum([WORK SHEET].[REJECTS]),'@@@@@@') " & _
                   "FROM [WORK SHEET],[MACHINE],[BARCODE] " & _
                   "WHERE  [WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
                          "[WORK SHEET].[OP_ID]      = [BARCODE].[OP_ID] AND " & _
                          "[WORK SHEET].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                          "[WORK SHEET].[MACHINE_ID] =" & MACHINE_ID & " AND " & _
                          "[WORK SHEET].[QUANTITY]<>0 " & _
                   "GROUP BY [WORK SHEET].[DATE_ID],[WORK SHEET].[MACHINE_ID],[WORK SHEET].[OP_ID]"
            
            '"format(60*sum([WORK SHEET].[QUANTITY])/sum([WORK SHEET].[TOTAL TIME]),'0,000') |Pcs/Hr" & _

            sSQLF = "    |^Date            |||"
            sSQLF = sSQLF & "|<Description                |<Operator                          |^         |>W.O.|>Hrs    "
            sSQLF = sSQLF & "|>Quantity        |>Defects  "

Else

        sSQL = "SELECT [WORK SHEET].[DATE_ID]                         AS [SQL 1]," & _
                         "[MACHINE].[MACHINE_ID]                   AS [SQL 2]," & _
                      "[WORK SHEET].[OP_ID]                        AS [SQL 3]," & _
                      "[BARCODE].[FIRST] &  ' ' & [BARCODE].[LAST] AS [SQL 4]," & _
                      "[WORK SHEET].[LOT NUM]," & _
                      "[WORK SHEET].[ATC PART]," & _
                      "format([WORK SHEET].[TOTAL TIME]/60,'0.0')," & _
                      "format([WORK SHEET].[QUANTITY],'###,##0')," & _
                      "format([WORK SHEET].[REJECTS],'##,##0') " & _
               "FROM [WORK SHEET],[MACHINE],[BARCODE] " & _
               "WHERE  [WORK SHEET].[MACHINE_ID] = [MACHINE].[MACHINE_ID] AND " & _
                      "[WORK SHEET].[OP_ID]      = [BARCODE].[OP_ID] AND " & _
                      "[WORK SHEET].[DATE_ID] BETWEEN #" & DATE_START_ID & "# AND #" & DATE_END_ID & "# AND " & _
                      "[WORK SHEET].[MACHINE_ID] =" & MACHINE_ID & " AND " & _
                      "[WORK SHEET].[QUANTITY]<>0 "
        
        '"format(60*sum([WORK SHEET].[QUANTITY])/sum([WORK SHEET].[TOTAL TIME]),'0,000') |Pcs/Hr" & _

        sSQLF = "    |^Date             ||"
        sSQLF = sSQLF & "|<Operator                         "
        sSQLF = sSQLF & "|^Lot W.O.                  |<ATC Part                         "
        sSQLF = sSQLF & "|>Hours   |>Quantity|>Defects"

End If

Data2.RecordSource = sSQL
Data2.Refresh
MSFlexGrid2.FormatString = sSQLF

End Sub

Private Sub cmdReset_Click()

DTPicker1.value = Date
DTPicker2.value = Date

optDay.value = True

cmdRefresh_Click

End Sub

Private Sub cmdWS_Click()

DATE_ID = DTPicker1.value

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    MACHINE = FR_Table.Fields("[MACHINE]")
    MACHINE_DESCRIPTION = FR_Table.Fields("[DESCRIPTION]")
    DEPT_ID = FR_Table.Fields("[DEPT_ID]")
End If

sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Operator = FR_Table.Fields("[FIRST]") & " " & FR_Table.Fields("[LAST]")
    SHIFT_ID = FR_Table.Fields("[SHIFT_ID]")
End If
FR_Table.Close
FR_Database.Close

Unload frmReviewWS

frmMain.Hide

frmWorkSheet1.Show

End Sub

Private Sub Form_Load()

Caption = "Review Work Sheets    " & ATC_DWG & "    " & ATC_VERSION

DTPicker1.value = Date
DTPicker2.value = Date

MSFlexGrid2.Top = 0
MSFlexGrid2.Width = 11000
MSFlexGrid2.Height = Me.Height - 800

Data2.DatabaseName = DB_OEE_WORKSHEET

cmdRefresh_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

frmOPScreen.Show

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = 2 And Shift = 1) Then
    If (fraTest.Enabled = True) Then
        fraTest.Enabled = False
        cmdDelete.Visible = True
    Else
        fraTest.Enabled = True
        cmdDelete.Visible = False
    End If
End If

End Sub

Private Sub MSFlexGrid2_Click()

If (MSFlexGrid2.Row = 0) Then
    Exit Sub
End If

MSFlexGrid2.Col = 1
DTPicker1.value = MSFlexGrid2.Text
DATE_ID = DTPicker1.value

MSFlexGrid2.Col = 2
MACHINE_ID = MSFlexGrid2.Text

MSFlexGrid2.Col = 3
OP_ID = MSFlexGrid2.Text

MSFlexGrid2.Col = 0
MSFlexGrid2.ColSel = MSFlexGrid2.Cols - 1

Set FR_Database = OpenDatabase(DB_OEE_WORKSHEET)

Dim sSQL As String
sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    lblMachineNo.Caption = FR_Table.Fields("[MACHINE]")
    lblDescription.Caption = FR_Table.Fields("[DESCRIPTION]")
End If

sSQL = "SELECT * FROM [BARCODE] WHERE [OP_ID]=" & OP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    lblOperator.Caption = FR_Table.Fields("[FIRST]") & " " & FR_Table.Fields("[LAST]")
End If

FR_Table.Close
FR_Database.Close

cmdWS.Enabled = True

End Sub

Private Sub optDay_Click()
cmdPrevious.Caption = "<< Day "
cmdNext.Caption = "Day >>"

DTPicker2.value = DTPicker1.value

cmdRefresh_Click
End Sub

Private Sub optWeek_Click()

DTPicker1.value = Format(DateAdd("d", -DTPicker1.DayOfWeek + 2, DTPicker1.value), "mm/dd/yyyy")
DTPicker2.value = Format(DateAdd("d", 6, DTPicker1.value), "mm/dd/yyyy")
 
cmdPrevious.Caption = "<< Week "
cmdNext.Caption = "Week >>"
cmdRefresh_Click
End Sub

Private Sub txtLot_GotFocus()
txtLot.SelStart = 0
txtLot.SelLength = Len(txtLot)
End Sub

Private Sub txtPart_GotFocus()
txtPart.SelStart = 0
txtPart.SelLength = Len(txtPart)
End Sub
