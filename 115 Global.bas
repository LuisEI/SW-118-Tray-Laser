Attribute VB_Name = "Global_Keyence"
'   File      : 115 Global.BAS
'   SW Engr   : Roger Soulagnet
'   DWG NO    : DWG NO 227-115 REV A
'   Date      : 06/05/2013
'   Program   : 115 Keyence Laser
'
'   ExecFile  : 115 Keyence Laser.EXE
'
'115 Keyence Laser 06/05/2013 SIGN_FACTOR Orientation
'115 Keyence Laser 09/24/2013 Production Release 1.0
'115 Keyence Laser 09/25/2013 No Chip Selected Error fix
'115 Keyence Laser 09/26/2013 CASE_ID "X" Extended E
'115 Keyence Laser 10/07/2013 CommandContinuousLaserGuide enable
'115 Keyence Laser 10/16/2013 Add Page to Fixture select on Worksheet and A/B Case
'115 Keyence Laser 10/30/2013 Add Text 2,3,4 Logo ATC,Omit
'115 Keyence Laser 01/24/2014 FROM [TBL SIZE LOC] Correction
'115 Keyence Laser 02/06/2014 Power Factors
'115 Keyence Laser 02/07/2014 MATRIX_ID is master
'115 Keyence Laser 02/10/2014 Data Base Backup
'115 Keyence Laser 02/18/2014 Text Font Spacing
'115 Keyence Laser 07/07/2014 Common Data Base Module
'115 Keyence Laser 03/05/2015 Corrections
'115 Keyence Laser 04/22/2015 Logo Top and ATC
'115 Keyence Laser 04/07/2015 fraLogo.Enabled = True
'115 Keyence Laser 08/24/2015
Option Explicit

Public Const ATC_DWG As String = "DWG NO 227-115 REV A"
Public Const ATC_VERSION As String = "08/24/2015"

Public Const TBL_ATC_DWG As String = "227-115"
Public Const TBL_NAME As String = "Keyence Laser"
Public Const TBL_EXECUTABLE As String = "115 Keyence Laser"


Public Const SIGN_FACTOR As Integer = -1
Public Const IN_TO_MM As Double = 25.4
Public Const MM_TO_IN As Double = 0.0393701

Public Const MAX_LASER_POWER As Integer = 100
Public Const MAX_SCAN_SPEED As Integer = 12000
Public Const MAX_FREQ As Integer = 400

Public EXTENDED_TEXT As Integer

Public RUN_MODE As Integer
Public MACHINE_TYPE_ID As Long

             
Public MATRIX_LID As Long
Public CASE_LID As String
             
Public INIT_RS232 As Integer
Public STOP_LASER As Integer
Public TRIGGER_MODE As Integer
             
Public Const DB_DPSS_LASER As String = "C:\ATC\115 MATRIX.MDB"

Public Const INCHES_TO_BITS As Long = 5842
Public Const INCHES_PER_SEC_TO_BITS_PER_SEC As Double = 5.842

Public TABLE_ID As Long             'UNIVERSAL TABLE ID
Public CONFIGURATION_ID As Long

Public ATC_PART_ID As String
Public SERIES_ID As String          '100A,200A,700A,100B,200B,700B
Public CASE_ID As String
Public TEXT_ID As String            'XXXT  ATC PART FORMAT AND TOLERANCE TO MARK

Public VALID_PART_ID As Long
Public VALID_LASER_ID As Long

Public TOLERANCE_ID As String

Public FIRE_COUNT_ID As Long

'========================================================================================
' KEYENECE LASER
'========================================================================================

Public K3_STRING As String
Public K3_STRING5 As String
Public K3_STRING6 As String
Public K3_STRING7 As String
Public K3_STRING_ATC As String

Public K3_LINE As String
Public K3_LOGO As String

Public STR_BLOCK_SHAPE As String
Public STR_BLOCK_TYPE As String

Public MARK_TEXT_FLAG As Long
Public MARK_TEXT_FLAG1 As Long
Public MARK_TEXT_FLAG2 As Long
Public MARK_TEXT_FLAG3 As Long
Public MARK_TEXT_FLAG4 As Long

Public MARK_LOGO_FLAG As Long
Public MARK_ATC_FLAG As Long
Public MARK_LINE_FLAG As Long
'====================================================================================
'   STRING PROPERTIES
'====================================================================================
Public STR_LASER_POWER As Double
Public STR_SCAN_SPEED As Long
Public STR_FREQ As Long

Public STR_X_POS As Double
Public STR_Y_POS As Double

Public TXT_X_POS2 As Double
Public TXT_Y_POS2 As Double
Public TXT_X_POS3 As Double
Public TXT_Y_POS3 As Double
Public TXT_X_POS4 As Double
Public TXT_Y_POS4 As Double

Public STR_CHAR_HT As Double
Public STR_CHAR_WD As Double
Public STR_CHAR_SP As Double

'====================================================================================
'   LINE PROPERTIES
'====================================================================================
Public LINE_LASER_POWER As Double
Public LINE_SCAN_SPEED As Long
Public LINE_FREQ As Long

Public LINE_START_X_POS As Double
Public LINE_START_Y_POS As Double
Public LINE_END_X_POS As Double
Public LINE_END_Y_POS As Double
Public LINE_Z_POS As Double

Public LINE_START_XB_POS As Double
Public LINE_START_YB_POS As Double
Public LINE_END_XB_POS As Double
Public LINE_END_YB_POS As Double
Public LINE_ZB_POS As Double

'FIRE POSITIONS
Public FR_LINE_START_X_POS As Double
Public FR_LINE_START_Y_POS As Double
Public FR_LINE_END_X_POS As Double
Public FR_LINE_END_Y_POS As Double

'====================================================================================
'   LOGO PROPERTIES
'====================================================================================
Public LOGO_LASER_POWER As Double
Public LOGO_SCAN_SPEED As Long
Public LOGO_FREQ As Long

Public LOGO_X_POS As Double
Public LOGO_Y_POS As Double
Public LOGO_WD As Double
Public LOGO_HT As Double

Public ATC_LOGO_X_POS As Double
Public ATC_LOGO_Y_POS As Double

Public BOARD_ID As Integer          'NATIONAL INSTRUMENTS MOTION CARD ID


Public STOP_DO As Integer       'STOP MOVE COMMAND MOTION

Public FR_Database As Database
Public FR_WorkSpace As Workspace
Public FR_Table As Recordset

Public TO_Database As Database
Public TO_WorkSpace As Workspace
Public TO_Table As Recordset

Public iInput(8) As Integer

Public NUM_PASSES As Integer
Public Y_OFFSET As Double
Public X_OFFSET As Double

Public Const LOGO_SIDE As Integer = 0
Public Const LOGO_TOP As Integer = 1
Public Const LOGO_ATC As Integer = 2
Public Const LOGO_OMIT As Integer = 3
Public LOGO_MODE As Integer

Public MARK_MODE As Integer
Public MARK_COUNT As Integer
 
'--------- CHIP LOCATION ARRAY [XLOC,YLOC,ANGLE]
Public Const XLOC As Integer = 0
Public Const YLOC As Integer = 1
Public Const Angle As Integer = 2
Public gdLocation(3, 100) As Double

Public Mark_Object(100) As Long

Type MATRIX_Type
        sCaption As String
        iRow As Integer
        iCol As Integer
        iCamRow As Integer  ' CAMERA ROW LOCATION
        iCamCol As Integer
        dXSpace As Double
        dYSpace As Double
End Type
Public Mat(60) As MATRIX_Type
Public MATRIX_ID As Long



Sub Main()

If App.PrevInstance Then
    End
End If
                                                        
Get_User
IP_ADDRESS = GetIPAddress
                            
               '971686002001
DEPT_ID = "LS"
MACHINE_TYPE = 0
MACHINE_DESCRIPTION = "Laser Keyence"
TEXT_ID = "NA"

Select Case Mid(IP_ADDRESS, 1, 8)
Case "10.0.38."
                    LOCATION_ID = "JR"
                    MACHINE_ID = 193
Case Else
                    LOCATION_ID = "NY"
                    MACHINE_ID = 193
End Select

ConfigComputer_DB (0)

Configuration (FREAD)

Select Case Mid(IP_ADDRESS, 1, 8)
Case "10.0.38."
                    LOCATION_ID = "JR"
                    MACHINE_ID = 193
                    DataBase_MODE = DATABASE_MODE_REM_JUAREZ
Case Else
                    LOCATION_ID = "NY"
                    MACHINE_ID = 193
                    DataBase_MODE = DATABASE_MODE_REM_NY
End Select

LOGO_MODE = LOGO_ATC

Configuration (FWRITE)

Select Case 0
Case 0
        MATRIX_ID = 31
        CASE_ID = "E"
Case 1
        MATRIX_ID = 35
        CASE_ID = "B"
End Select

DataBase_Address

Get_Matrix_Parameters

Select Case 1
Case 0
        frmKey.Hide
        frmMain.Show                '[Main Keyence Laser Matrix Form]
Case 1
        frmKey.Hide
        frmOPScreen.Show            '[Keyence Laser Production Screen] PRODUCTION MODE
Case 2
        frmKey.Show
Case 3
        frmPowerFactors.Show
Case 4
        frmMatrix.Show              '[Laser Fixture Select]
Case 5
        frmReviewWSNew.Show            '[Review Work Sheets]
Case 6
        frmConfiguration.Show       '[Configuration Dual DPSS Laser]
Case 7
        frmMessage.Show             '[Message Alert eg Check Out Bowl Feed]
Case 8
        frmIO.Show                  '[PCI I/O Board Test]
End Select

End Sub

'
'   FILE : 115 CONFIG.TXT
'
Sub Configuration(iMode As Integer)

On Error GoTo Error_Handler

Dim iFilenum As Integer
Dim sFilename As String, sTemp As String
 
sFilename = "C:\ATC\115 CONFIG.TXT"
iFilenum = FreeFile
If (Len(Dir$(sFilename)) = 0) Then
    iMode = FWRITE
    DataBase_MODE = DATABASE_MODE_LCL
    LOCATION_ID = "JR"
    MATRIX_ID = 31
    CASE_ID = "E"
    RUN_MODE = 0
    INIT_RS232 = 0
    EXTENDED_TEXT = 0
End If

Select Case iMode
Case FREAD
            Open sFilename For Input As iFilenum
            Input #iFilenum, sTemp: DataBase_MODE = InfoVal(sTemp)
            Input #iFilenum, sTemp: LOCATION_ID = InfoStr(sTemp)
            Input #iFilenum, sTemp: MATRIX_ID = InfoVal(sTemp)
            Input #iFilenum, sTemp: CASE_ID = InfoStr(sTemp)
            Input #iFilenum, sTemp: RUN_MODE = InfoVal(sTemp)
            Input #iFilenum, sTemp: INIT_RS232 = InfoVal(sTemp)
            Input #iFilenum, sTemp: EXTENDED_TEXT = InfoVal(sTemp)
Case FWRITE
            Open sFilename For Output As #iFilenum
            Print #iFilenum, "[01] DATABASE_MODE [0:1:2:4]     ="; DataBase_MODE
            Print #iFilenum, "[02] LOCATION_ID   [NY:JR]       ="; LOCATION_ID
            Print #iFilenum, "[03] MATRIX_ID                   ="; MATRIX_ID
            Print #iFilenum, "[04] CASE_ID                     ="; CASE_ID
            Print #iFilenum, "[05] RUN_MODE [Test:0 Prod:1]    ="; RUN_MODE
            Print #iFilenum, "[06] INIT_RS232  0:Disabled      ="; INIT_RS232
            Print #iFilenum, "[07] EXTENDED_TEXT               ="; EXTENDED_TEXT
            Print #iFilenum, "[08] MACHINE_ID                  ="; MACHINE_ID
End Select
Close iFilenum

Exit Sub

Error_Handler:

Resume Next

End Sub

Public Sub DataBase_Address()

'================================== WO SCHED MASTER
Select Case DataBase_MODE
Case DATABASE_MODE_LCL
                DB_MASTER_SCHEDULE = "C:\ATC\WO SCHED MASTER.mdb"
Case DATABASE_MODE_REM_NY
                DB_MASTER_SCHEDULE = SERVER_DB_NY & "WO SCHED MASTER.mdb"
Case DATABASE_MODE_FIL
                DB_MASTER_SCHEDULE = SERVER_ADDR_OEE & "WO SCHED MASTER.mdb"
Case DATABASE_MODE_REM_JUAREZ
                DB_MASTER_SCHEDULE = "C:\ATC\WO SCHED MASTER.mdb"
End Select

'================================== OEE WORK SHEET
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY
                DB_OEE_WORKSHEET = SERVER_DB_NY & "OEE SPM JR.MDB"
Case DATABASE_MODE_LCL
               DB_OEE_WORKSHEET = "C:\ATC\OEE SPM JR.MDB"
Case DATABASE_MODE_FIL
                
Case DATABASE_MODE_REM_JUAREZ
                DB_OEE_WORKSHEET = SERVER_DB_JR & "OEE SPM JR.MDB"                  ' New Test
End Select


End Sub

Public Sub Get_Matrix_Parameters()

Dim sSQL As String
Set FR_Database = OpenDatabase(DB_DPSS_LASER)

sSQL = "SELECT * FROM [FIXTURE] WHERE [MATRIX ID] =" & MATRIX_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then

    X_OFFSET = FR_Table.Fields("[FIX X LOC]")
    Y_OFFSET = FR_Table.Fields("[FIX Y LOC]")

    STR_LASER_POWER = FR_Table.Fields("[LaserPower]")
    STR_SCAN_SPEED = FR_Table.Fields("[Markspeed]")
    STR_FREQ = FR_Table.Fields("[Frequency]")
    
    LINE_LASER_POWER = FR_Table.Fields("[LaserPower 2]")
    LINE_SCAN_SPEED = FR_Table.Fields("[Markspeed 2]")
    LINE_FREQ = FR_Table.Fields("[Frequency 2]")
                
    LOGO_LASER_POWER = FR_Table.Fields("[LaserPower 3]")
    LOGO_SCAN_SPEED = FR_Table.Fields("[Markspeed 3]")
    LOGO_FREQ = FR_Table.Fields("[Frequency 3]")
        
    LINE_START_X_POS = FR_Table.Fields("[A LOC X]")
    LINE_START_Y_POS = FR_Table.Fields("[A LOC Y]")
    LINE_END_X_POS = FR_Table.Fields("[A LOC X]") + FR_Table.Fields("[A WIDTH]")
    LINE_END_Y_POS = FR_Table.Fields("[A LOC Y]")
    LINE_Z_POS = FR_Table.Fields("[A LOC Z]")

    LINE_START_XB_POS = FR_Table.Fields("[B LOC X]")
    LINE_START_YB_POS = FR_Table.Fields("[B LOC Y]")
    LINE_END_XB_POS = FR_Table.Fields("[B LOC X]") + FR_Table.Fields("[B WIDTH]")
    LINE_END_YB_POS = FR_Table.Fields("[B LOC Y]")
    LINE_ZB_POS = FR_Table.Fields("[B LOC Z]")
    
    NUM_PASSES = FR_Table.Fields("[NUM PASSES]")

End If

sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [CASE SIZE] ='" & CASE_ID & "'"

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then

    Select Case LOGO_MODE
    Case LOGO_SIDE
                LOGO_X_POS = FR_Table.Fields("[LOGO LX SIDE]")
                LOGO_Y_POS = FR_Table.Fields("[LOGO LY SIDE]")
    Case Else
                LOGO_X_POS = FR_Table.Fields("[LOGO LX TOP]")
                LOGO_Y_POS = FR_Table.Fields("[LOGO LY TOP]")
    End Select

    ATC_LOGO_X_POS = FR_Table.Fields("[ATC  X]") * IN_TO_MM
    ATC_LOGO_Y_POS = FR_Table.Fields("[ATC  Y]") * IN_TO_MM

    STR_CHAR_HT = FR_Table.Fields("[TEXT YSIZE]") * IN_TO_MM
    STR_CHAR_WD = FR_Table.Fields("[TEXT XSIZE]")

    STR_CHAR_SP = FR_Table.Fields("[TEXT SPACE]")

    LOGO_WD = FR_Table.Fields("[LOGO XSIZE]") * IN_TO_MM
    LOGO_HT = FR_Table.Fields("[LOGO YSIZE]") * IN_TO_MM

    TXT_X_POS2 = FR_Table.Fields("[TXT2 X]") * IN_TO_MM
    TXT_Y_POS2 = FR_Table.Fields("[TXT2 Y]") * IN_TO_MM
    
    TXT_X_POS3 = FR_Table.Fields("[TXT3 X]") * IN_TO_MM
    TXT_Y_POS3 = FR_Table.Fields("[TXT3 Y]") * IN_TO_MM
    
    TXT_X_POS4 = FR_Table.Fields("[TXT4 X]") * IN_TO_MM
    TXT_Y_POS4 = FR_Table.Fields("[TXT4 Y]") * IN_TO_MM


End If

'115 Keyence Laser 01/14/2014 FROM [TBL SIZE LOC] Correction

sSQL = "SELECT * FROM [TBL SIZE LOC] "

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        FR_Table.Edit
        Select Case FR_Table.Fields("[SIZE_LOC_ID]")
        Case 1
                FR_Table.Fields("[CASE SIZE]") = "B"
                FR_Table.Fields("[CASE NAME]") = "B"
        Case 2
                FR_Table.Fields("[CASE SIZE]") = "C"
                FR_Table.Fields("[CASE NAME]") = "C"
        Case 3
                FR_Table.Fields("[CASE SIZE]") = "E"
                FR_Table.Fields("[CASE NAME]") = "E Normal"
        Case 4
                FR_Table.Fields("[CASE SIZE]") = "X"
                FR_Table.Fields("[CASE NAME]") = "E Extended"
        Case 5
                FR_Table.Fields("[CASE SIZE]") = "A"
                FR_Table.Fields("[CASE NAME]") = "A"
        Case 6
                FR_Table.Fields("[CASE SIZE]") = "R"
                FR_Table.Fields("[CASE NAME]") = "R"
        Case Else
        End Select
        FR_Table.Update
        FR_Table.MoveNext
    Loop
End If
FR_Table.Close
FR_Database.Close


End Sub
