Attribute VB_Name = "GLOBAL"
'   File      : 118 GLOBAL.BAS
'   SW Engr   : Roger Soulagnet
'   DWG NO    : DWG NO 227-118 REV A
'   ECN NO    :
'   Date      : 01/28/2015
'   Program   : 118 Tray Laser
'   Exe File  : 118 Tray Laser.EXE

'   Input Files  :
'
'---------------------------------------------------------------------------------
'118 Tray Laser chg 01/28/2015 Start
'118 Tray Laser chg 06/08/2015 Logo Correction
'118 Tray Laser chg 06/17/2015 Logo Default and Fire commandbutton no Tab
'118 Tray Laser chg 06/23/2015 Select Fixture Scan Work Order
'118 Tray Laser chg 06/30/2015 Corrections
'118 Tray Laser chg 07/14/2016 H Case 301-476
'118 Tray Laser chg 07/27/2016 ABRASIVE LINES
'118 Tray Laser chg 08/02/2016 Rotate
'118 Tray Laser chg 08/31/2016 [14]  301-H89  E Case 6 X 4 Vertical
'118 Tray Laser chg 09/14/2016 TRAY_MARK_ANGLE
'118 Tray Laser chg 09/14/2016 SIZE_LOC_ID
'118 Tray Laser chg 10/21/2016 [MARK PARA]
'118 Tray Laser chg 10/31/2016 [CAMERA POS]
'118 Tray Laser chg 12/20/2016 Calibration Laser Height
'118 Tray Laser chg 05/04/2017 Abrasize Camera Position and X,Y Offsets
'118 Tray Laser chg 02/27/2019 10X10 CASES S,L,F

Option Explicit

Public Const ATC_DWG As String = "DWG NO 227-118"
Public Const ATC_VERSION As String = "03/20/2019"

Public Const TBL_ATC_DWG As String = "227-118"
Public Const TBL_NAME As String = "DPSS Tray Laser"
Public Const TBL_EXECUTABLE As String = "118 Tray Laser"

Public ATC_LASER_BD As String
                                                       
Public Const CAL_STICK As Double = 10.235
                                                       
Public Const MARK_ANGLE_DEFAULT As Integer = 0
Public Const MARK_ANGLE_ROTATED As Integer = 1
Public MARK_ANGLE As Integer
                                                      
Public TRAY_MARK_ANGLE As Integer
                                                      
Public FIRE_COUNT_ID As Long
Public FIRE_COUNT As Long

Public INITIALIZE_TRAY As Integer

Public Const WORKSHEET_MODE As Integer = 1
Public Const TEST_MODE As Integer = 0
Public OP_MODE As Integer

Public iInput(8) As Integer

Public StartTime As Double

Public COUNT_ID As Long

Public FORM_LOC_X As Long
Public FORM_LOC_Y As Long

Public Const INCHES_TO_BITS As Long = 5842
Public Const INCHES_PER_SEC_TO_BITS_PER_SEC As Double = 5.842

Public PATH_SCHEDULE_LCL As String
Public PATH_SCHEDULE_REM As String
Public PATH_SCHEDULE_FIL As String

Public PATH_OEE_DB_LCL As String
Public PATH_OEE_DB_REM As String
Public PATH_OEE_DB_FIL As String

Public OPTION1_ENABLE As Integer
Public OPTION2_ENABLE As Integer
Public OPTION3_ENABLE As Integer

'=================================================================================
'Public CASE_ID As String            ' not used on Tray forms only Main and Power
Public VALID_PART_ID As Long
Public TOLERANCE_ID As String
Public TEXT_ID As String            'XXXT  ATC PART FORMAT AND TOLERANCE TO MARK

'=================================================================================

Public BOARD_ID As Integer
Public AXIS_ID As Integer

Public ACCEL_ID As Double
Public DECEL_ID As Double
Public VELOCITY_ID As Double

Public ObjectCount As Long

Public ObjIndex As Long
Public ProfileIndex As Long

Public Markspeed As Double
Public Jumpspeed As Double

Public Markspeed_Bits As Long
Public Jumpspeed_Bits As Long

Public Jumpdelay As Long
Public Markdelay As Long
Public Polygondelay As Long
Public Laserpower As Single
Public Laseroffdelay As Long
Public Laserondelay As Long
Public Eightbitword As Long
Public T1 As Double
Public T2 As Long
Public zAxis As Long
Public Varijumpdelay As Long
Public Varijumplength As Long
Public Wobblesize As Long
Public Wobblefrequency As Double
Public Powerreset As Long
Public Varipolydelay As Long

Public BusyFlag As Long
Public BusyFlag2 As Long

Public VALID_LASER_ID As Long
'Public ATC_DWG_ID As String
Public Load_Job As Integer

Public TRAY_ID As Long
Public POWER_ID As Long
Public SIZE_LOC_ID As Long

Public LASER_TXT1 As String
Public LASER_TXT2 As String
Public LASER_TXT3 As String
Public LASER_TXT4 As String

Public TRAY_Y_OFFSET As Double
Public TRAY_X_OFFSET As Double

Public LICA_XOFF As Double
Public LICA_YOFF As Double

Public SEG_Y_DIST As Double
Public SEG_Y_DIST_0 As Double
Public SEG_Y_DIST_1 As Double

Public Const LOGO_SIDE As Integer = 0
Public Const LOGO_TOP As Integer = 1
Public Const LOGO_ATC As Integer = 2
Public Const LOGO_OMIT As Integer = 3
Public Const LOGO_ABRASIVE As Integer = 4

Public LOGO_MODE As Integer

Public REP_ABRASIVE As Integer

Public MARK_MODE As Integer
Public MARK_COUNT As Integer
 
'--------- CHIP LOCATION ARRAY [XLOC,YLOC,ANGLE]
Public Const XLOC As Integer = 0
Public Const YLOC As Integer = 1
Public Const Angle As Integer = 2
Public gdLocation(3, 100) As Double

Public ObjHPosition(20) As Double
Public ObjVPosition(20) As Double

Public FIRE_MATRIX(12, 100) As Integer
Public SEGMENT_ID As Integer
Public SEGMENTS_SELECT As Integer

Public MATRIX_ID As Long

Public FR_Database As Database
Public FR_WorkSpace As Workspace
Public FR_Table As Recordset
 
Public TO_Database As Database
Public TO_WorkSpace As Workspace
Public TO_Table As Recordset
 
Public lCount As Long
'
'   SERIALIZATION
'
Public SERIAL_START_NUMBER As Integer

'Public AutomationInterface As New winlase.Automate
'Public LecInterface As New winlase.Lec


Sub Main()
              
    If App.PrevInstance Then
        End
    End If
                                                
    Get_User
    IP_ADDRESS = GetIPAddress
    
    Select Case Mid(IP_ADDRESS, 1, 8)
    Case "10.0.38."
                        LOCATION_ID = "JR"
                        DataBase_MODE = DATABASE_MODE_REM_JUAREZ
    Case Else
                        LOCATION_ID = "NY"
                        DataBase_MODE = DATABASE_MODE_REM_NY
    End Select
    
    Configuration (FREAD)
    
    LOCATION_ID = "JR"
    
    Update_2_DataBases
    
    DataBase_Address
    
    LoveLetter
                  
    ConfigComputer_DB (0)

    OP_MODE = 1 '0 NORMAL 1 TEST
   ' MACHINE_ID = 213
    'MACHINE_NUMBER = 6
    
    DEPT_ID = "LS"
    MACHINE_TYPE = 9
    'MACHINE_DESCRIPTION = "Tray Laser"
        
    Set FR_Database = OpenDatabase(DB_ELECT_OP_MACHINE)
    Dim sSQL As String
    sSQL = "SELECT * FROM [MACHINE] WHERE [MACHINE_ID]=" & MACHINE_ID
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    
    If (FR_Table.RecordCount <> 0) Then
        MACHINE_NUMBER = FR_Table.Fields("[MACHINE]")
        MACHINE_DESCRIPTION = FR_Table.Fields("[DESCRIPTION]")
    End If
    
    FR_Table.Close
    FR_Database.Close
  
    Configuration (FWRITE)
        
    Select Case 0
    Case 0
            Select Case OP_MODE
            Case 0
                    frmMain.Show
            Case 1
                    frmOPScreen.Show
            End Select
    Case 1
            CASE_ID = "E"
            SIZE_LOC_ID = 3
            POWER_ID = 1
            TRAY_ID = 5
            frmPowerFactors.Show
    Case 1
            frmOPScreen.Show
    Case 2
            frmMain.Show        'Tray Laser Main
    Case 3
            frmMotion.Show      'NI Motion Functions
    Case 4
            frmTray.Show        'Tray Configuration
    Case 5
            frmConfiguration.Show
    Case 13
            
            TRAY_ID = 13
            POWER_ID = 15
            SIZE_LOC_ID = 10
            LOGO_MODE = LOGO_ABRASIVE
            
            frm103.Show         'ATC 301-476           'H' Case Carrier Tray
    Case 6
            TRAY_ID = 5
            frm412.Show         'ATC 301-412           'E' Case Carrier Tray
    Case 7
            TRAY_ID = 1
            frm413.Show         'ATC 301-413           'E' Encapsulated Case Carrier Tray
    Case 8
            TRAY_ID = 12
            frm414.Show         'Tray 414 C Case
    Case 9
            TRAY_ID = 8
            frm10x10.Show
    Case 10
            TRAY_ID = 9
            SIZE_LOC_ID = 1
            POWER_ID = 76
            frm20x20.Show       'B' Case 10 X 10 Carrier Tray
    End Select
       
End Sub
 
 
Public Function BinaryToDecimal(Binary As String) As Long
Dim n As Long
Dim s As Integer

    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ _
            (s - 1)))
    Next s

    BinaryToDecimal = n
End Function

 
Sub Configuration(iMode As Integer)

Dim iFilenum As Integer, I As Integer
Dim sFileName As String, sTemp As String
 
sFileName = "C:\ATC\118 Configuration New.TXT"
iFilenum = FreeFile
If (Len(Dir$(sFileName)) = 0) Then
    iMode = FWRITE
    BOARD_ID = 1
    LOCATION_ID = "JR"
    DataBase_MODE = DATABASE_MODE_REM_JUAREZ
    MACHINE_ID = 213
    INITIALIZE_TRAY = 0
    OP_MODE = 0
End If

Select Case iMode
Case FREAD
            Open sFileName For Input As iFilenum
            Input #iFilenum, sTemp: OP_MODE = InfoVal(sTemp)
            Input #iFilenum, sTemp: INITIALIZE_TRAY = InfoVal(sTemp)
            Input #iFilenum, sTemp: LOCATION_ID = InfoStr(sTemp)
            Input #iFilenum, sTemp: BOARD_ID = InfoVal(sTemp)
            Input #iFilenum, sTemp: DataBase_MODE = InfoVal(sTemp)
            Input #iFilenum, sTemp: FORM_LOC_X = InfoVal(sTemp)
            Input #iFilenum, sTemp: FORM_LOC_Y = InfoVal(sTemp)
            Input #iFilenum, sTemp: MACHINE_ID = InfoVal(sTemp)
            
Case FWRITE
            Open sFileName For Output As #iFilenum
            Print #iFilenum, "(01) OP_MODE Test:0 WS:1     ="; OP_MODE
            Print #iFilenum, "(02) INITIALIZE_TRAY         ="; INITIALIZE_TRAY
            Print #iFilenum, "(03) LOCATION_ID             ="; LOCATION_ID
            Print #iFilenum, "(04) I/O Board No            ="; BOARD_ID
            Print #iFilenum, "(05) DATABASE_MODE           ="; DataBase_MODE
            Print #iFilenum, "(06) FORM_LOC_X              ="; FORM_LOC_X
            Print #iFilenum, "(07) FORM_LOC_Y              ="; FORM_LOC_Y
            Print #iFilenum, "(08) MACHINE_ID              ="; MACHINE_ID
            Print #iFilenum, "(09) Spare                   ="; 0
            Print #iFilenum, "(10) Spare                   ="; 0
            Print #iFilenum, "(11) Spare                   ="; 0
End Select
Close iFilenum

End Sub

Public Sub DataBase_Address()

'================================== ATC Electrical Test Parameters
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY
                DB_ATC_Electrical_Test = SERVER_DB_NY & "ATC Electrical Test.MDB"
Case DATABASE_MODE_LCL
                DB_ATC_Electrical_Test = "C:\ATC\ATC Electrical Test.MDB"
Case DATABASE_MODE_FIL
                DB_ATC_Electrical_Test = SERVER_ADDR_SPC
Case DATABASE_MODE_REM_JUAREZ
                DB_ATC_Electrical_Test = "C:\ATC\ATC Electrical Test.MDB"
End Select

'================================== WO SCHED MASTER
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY
                DB_MASTER_SCHEDULE = SERVER_DB_NY & "WO SCHED MASTER.mdb"
Case DATABASE_MODE_LCL
                DB_MASTER_SCHEDULE = "C:\ATC\WO SCHED MASTER.mdb"
Case DATABASE_MODE_FIL
                DB_MASTER_SCHEDULE = SERVER_ADDR_MAS
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
                DB_OEE_WORKSHEET = SERVER_ADDR_OEE
Case DATABASE_MODE_REM_JUAREZ
                DB_OEE_WORKSHEET = SERVER_DB_JR & "OEE SPM JR.MDB"                  ' New Test
End Select

'================================== OEE WORK SHEET

Select Case MACHINE_ID
Case 213
        Select Case DataBase_MODE
        Case DATABASE_MODE_REM_NY
                        ATC_LASER_BD = SERVER_DB_NY & "118 LASER MATRIX.MDB"
        Case DATABASE_MODE_LCL
                        ATC_LASER_BD = "C:\ATC\118 LASER MATRIX.MDB"
        Case DATABASE_MODE_FIL
                         
        Case DATABASE_MODE_REM_JUAREZ
                        ATC_LASER_BD = SERVER_DB_JR & "118 LASER MATRIX.MDB"                  ' New Test
        End Select
Case Else
        Select Case DataBase_MODE
        Case DATABASE_MODE_REM_NY
                        ATC_LASER_BD = SERVER_DB_NY & "118 LASER MATRIXB.MDB"
        Case DATABASE_MODE_LCL
                        ATC_LASER_BD = "C:\ATC\118 LASER MATRIXB.MDB"
        Case DATABASE_MODE_FIL
                         
        Case DATABASE_MODE_REM_JUAREZ
                        ATC_LASER_BD = SERVER_DB_JR & "118 LASER MATRIXB.MDB"                  ' New Test
        End Select
End Select

'================================== DB_ELECT_OP_MACHINE
Select Case DataBase_MODE
Case DATABASE_MODE_REM_NY
                DB_ELECT_OP_MACHINE = "C:\ATC\ATC Electrical Test.MDB"
Case DATABASE_MODE_LCL
                DB_ELECT_OP_MACHINE = "C:\ATC\ATC Electrical Test.MDB"
Case DATABASE_MODE_REM_JUAREZ
                DB_ELECT_OP_MACHINE = "C:\ATC\ATC Electrical Test.MDB"
End Select

End Sub


Public Function Get_Title() As String

Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String
sSQL = "SELECT [TRAY_ID],[CASE],[TITLE],[ROWS],[COLS],[ATC DWG]," & _
                "[Segments],[Spacing]," & _
                "[L X OffSet],[L Y OffSet]," & _
                "[X 0],[X 1],[X 2],[X 3],[X 4],[X 5],[PAGE],[Y OffSet A],[ROTATION] " & _
        "FROM [TBL Tray Config] WHERE [TRAY_ID] = " & TRAY_ID
                 
Set FR_Table = FR_Database.OpenRecordset(sSQL)
     
Get_Title = FR_Table.Fields("[ATC DWG]") & "  " & FR_Table.Fields("[ROWS]") & " X " & FR_Table.Fields("[COLS]") & "  " & FR_Table.Fields("[CASE]") & " Case    " & FR_Table.Fields("[TITLE]")

FR_Database.Close

End Function

Public Sub GlobalUpdate()

Dim MARK_PARAMETER As String

Dim Markspeed As Double
Dim frequency As Double
Dim PulseWidth As Double

Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String

sSQL = "SELECT * FROM [TBL Power] WHERE [TBL_ID] =" & POWER_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
MARK_PARAMETER = FR_Table.Fields("[MARK PARA]")
Markspeed = FR_Table.Fields("[Markspeed]")
frequency = FR_Table.Fields("[Frequency]")
PulseWidth = FR_Table.Fields("[PulseWidth]")
 
sSQL = "SELECT * FROM [TBL Power] WHERE [MARK PARA] =" & MARK_PARAMETER

Set FR_Table = FR_Database.OpenRecordset(sSQL)
 Dim COUNT As Integer
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
            FR_Table.Edit
            FR_Table.Fields("[Frequency]") = frequency
            FR_Table.Fields("[Markspeed]") = Markspeed
            FR_Table.Fields("[PulseWidth]") = PulseWidth
            FR_Table.Update
            FR_Table.MoveNext
            COUNT = COUNT + 1
    Loop
End If
FR_Table.Close
FR_Database.Close

MsgBox "Complete COUNT" & COUNT, vbInformation, "ATC Data Base System"

End Sub


Public Sub GlobalHeightUpdate()

Dim RATIO_STEPS_PER_INCH As Double

Dim sSQL As String

sSQL = "SELECT * FROM [TBL Power] WHERE [TBL_ID] =" & POWER_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
                        '[ZHEIGHT] STEPS PER  [HEIGHT] - INCHES
 
RATIO_STEPS_PER_INCH = FR_Table.Fields("[ZHEIGHT]") / (CAL_STICK + FR_Table.Fields("[HEIGHT]"))
  
sSQL = "SELECT * FROM [TBL Power] WHERE  [ACTIVE] = Yes  ORDER BY [ORDER]"

Set FR_Table = FR_Database.OpenRecordset(sSQL)
 
Dim COUNT As Integer
 
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
            FR_Table.Edit
            FR_Table.Fields("[ZHEIGHT]") = (CAL_STICK + FR_Table.Fields("[HEIGHT]")) * RATIO_STEPS_PER_INCH
            FR_Table.Update
            FR_Table.MoveNext
             COUNT = COUNT + 1
    Loop
End If
FR_Table.Close
FR_Database.Close

MsgBox "Complete COUNT" & COUNT, vbInformation, "ATC Data Base System"

End Sub
