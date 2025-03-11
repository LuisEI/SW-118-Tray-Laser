Attribute VB_Name = "Laser_Module"
Option Explicit
'
'           GREATER THAN 10 FIRED
'               Date,Time, Matrix
Sub WorkLog()
    
Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String
sSQL = "SELECT * FROM [WORK LOG] WHERE [FIX_ID]=" & MATRIX_ID & " AND [DATE_ID]=" & Date

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount = 0) Then
    FR_Table.AddNew
Else
    FR_Table.Edit
End If

FR_Table.Fields("[FIX_ID]") = MATRIX_ID
FR_Table.Fields("[DATE_ID]") = Date
FR_Table.Fields("[QUANTITY]") = FR_Table.Fields("[QUANTITY]") + MARK_COUNT
 
FR_Table.Update

FR_Table.Close
FR_Database.Close
       
End Sub

Public Sub Initialize_Fire_Matrix()

Dim i, j As Integer
For i = 0 To 10
    For j = 0 To 100
        FIRE_MATRIX(i, j) = 0
    Next j
Next i

End Sub

Public Sub LoadDPSS()

'=============================================
'[1] SetObjCharString
'=============================================
Select Case MARK_MODE
Case 1
        AutomationInterface.SetObjCharString 0, LASER_TXT1
Case 2
        AutomationInterface.SetObjCharString 0, LASER_TXT1
        AutomationInterface.SetObjCharString 1, LASER_TXT2
Case 3
        AutomationInterface.SetObjCharString 0, LASER_TXT1
        AutomationInterface.SetObjCharString 1, LASER_TXT2
        AutomationInterface.SetObjCharString 2, LASER_TXT3
Case 4
        AutomationInterface.SetObjCharString 0, LASER_TXT1
        AutomationInterface.SetObjCharString 1, LASER_TXT2
        AutomationInterface.SetObjCharString 2, LASER_TXT3
        AutomationInterface.SetObjCharString 3, LASER_TXT4
End Select

'=============================================
'[2] SetObjSize
'=============================================

Dim sSQL As String
Set FR_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)

Dim HTextSize As Long
Dim VTextSize As Long

HTextSize = Format(FR_Table.Fields("[TEXT XSIZE]") * INCHES_TO_BITS, "0")
VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS, "0")

'118 Tray Laser chg 07/19/2018 SIZE_LOC_ID
Select Case SIZE_LOC_ID
Case 100


Case Else
        AutomationInterface.SetObjSize 0, HTextSize, VTextSize
        AutomationInterface.SetObjSize 1, HTextSize, VTextSize
        AutomationInterface.SetObjSize 2, HTextSize, VTextSize
        AutomationInterface.SetObjSize 3, HTextSize, VTextSize
End Select

Select Case LOGO_MODE
Case LOGO_SIDE, LOGO_TOP
                HTextSize = Format(FR_Table.Fields("[LOGO XSIZE]") * INCHES_TO_BITS, "0")
                VTextSize = Format(FR_Table.Fields("[LOGO YSIZE]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 4, HTextSize, VTextSize
Case LOGO_OMIT

Case LOGO_ATC
                HTextSize = Format(FR_Table.Fields("[TEXT XSIZE]") * INCHES_TO_BITS, "0")
                VTextSize = Format(FR_Table.Fields("[TEXT YSIZE]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 5, HTextSize, VTextSize
                
Case LOGO_ABRASIVE

                'SIZE = LENGTH     HORIZONTAL OR VERTICAL LINE
                
                HTextSize = Format(FR_Table.Fields("[LEN H]") * INCHES_TO_BITS, "0")
                VTextSize = 0
                AutomationInterface.SetObjSize 6, HTextSize, VTextSize
                AutomationInterface.SetObjSize 7, HTextSize, VTextSize

                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                HTextSize = 0
                AutomationInterface.SetObjSize 8, HTextSize, VTextSize

                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 9, HTextSize, VTextSize

                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 10, HTextSize, VTextSize
                                
                VTextSize = Format(FR_Table.Fields("[LEN V]") * INCHES_TO_BITS, "0")
                AutomationInterface.SetObjSize 11, HTextSize, VTextSize
End Select

FR_Database.Close

'MsgBox "SetObjSize Complete", vbInformation, "Laser"

Set FR_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT * FROM [TBL POWER] WHERE [TBL_ID] =" & POWER_ID

Set FR_Table = FR_Database.OpenRecordset(sSQL)

ProfileIndex = 0

Jumpspeed = FR_Table.Fields("[Jumpspeed]")
Jumpspeed_Bits = Format(Jumpspeed * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")

Jumpdelay = FR_Table.Fields("[Jumpdelay]")
Polygondelay = FR_Table.Fields("[Polygondelay]")
Markdelay = FR_Table.Fields("[Markdelay]")
Laserpower = FR_Table.Fields("[Laserpower]")
Laseroffdelay = FR_Table.Fields("[Laseroffdelay]")
Laserondelay = FR_Table.Fields("[Laserondelay]")
Wobblesize = FR_Table.Fields("[Wobblesize]")
Wobblefrequency = FR_Table.Fields("[Wobblefrequency]")

Markspeed = FR_Table.Fields("[Markspeed]")
Markspeed_Bits = Format(Markspeed * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")

T1 = FR_Table.Fields("[Frequency]")
T2 = FR_Table.Fields("[PulseWidth]")

Select Case LOGO_MODE
Case LOGO_ABRASIVE
    Markspeed = FR_Table.Fields("[ABRASIZE_Markspeed]")
    Markspeed_Bits = Format(Markspeed * INCHES_PER_SEC_TO_BITS_PER_SEC, "0")
    T1 = FR_Table.Fields("[ABRASIZE_Frequency]")
    T2 = FR_Table.Fields("[ABRASIZE_PulseWidth]")
End Select

Dim i As Integer
For i = 0 To ObjectCount - 1
        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        AutomationInterface.SetObjProfile i, ProfileIndex, Markspeed_Bits, Jumpspeed_Bits, Jumpdelay, Markdelay, Polygondelay, Laserpower, Laseroffdelay, Laserondelay, Eightbitword, T1, T2, zAxis, Varijumpdelay, Varijumplength, Wobblesize, Wobblefrequency, Powerreset, Varipolydelay
Next i

sSQL = "SELECT * FROM [TBL SIZE LOC] WHERE [SIZE_LOC_ID] =" & SIZE_LOC_ID
Set FR_Table = FR_Database.OpenRecordset(sSQL)
  
ObjHPosition(0) = FR_Table.Fields("[TXT1 X]")
ObjVPosition(0) = FR_Table.Fields("[TXT1 Y]")
ObjHPosition(1) = FR_Table.Fields("[TXT2 X]")
ObjVPosition(1) = FR_Table.Fields("[TXT2 Y]")
ObjHPosition(2) = FR_Table.Fields("[TXT3 X]")
ObjVPosition(2) = FR_Table.Fields("[TXT3 Y]")
ObjHPosition(3) = FR_Table.Fields("[TXT4 X]")
ObjVPosition(3) = FR_Table.Fields("[TXT4 Y]")
   
Select Case LOGO_MODE
Case LOGO_SIDE
    ObjHPosition(4) = FR_Table.Fields("[LOGO LX SIDE]")
    ObjVPosition(4) = FR_Table.Fields("[LOGO LY SIDE]")
Case LOGO_TOP
    ObjHPosition(4) = FR_Table.Fields("[LOGO LX TOP]")
    ObjVPosition(4) = FR_Table.Fields("[LOGO LY TOP]")
Case LOGO_ATC
    ObjHPosition(5) = FR_Table.Fields("[ATC  X]")
    ObjVPosition(5) = FR_Table.Fields("[ATC  Y]")
Case LOGO_ABRASIVE
    'ABRASIZE
    ObjHPosition(6) = FR_Table.Fields("[LINE X1]")
    ObjVPosition(6) = FR_Table.Fields("[LINE Y1]")
    ObjHPosition(7) = FR_Table.Fields("[LINE X2]")
    ObjVPosition(7) = FR_Table.Fields("[LINE Y2]")
        
    ObjHPosition(8) = FR_Table.Fields("[LINE X1]") + (FR_Table.Fields("[LEN H]") * 0.5) - (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(8) = FR_Table.Fields("[LINE Y1]")
    ObjHPosition(9) = FR_Table.Fields("[LINE X1]") + (FR_Table.Fields("[LEN H]") * 0.5) + (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(9) = FR_Table.Fields("[LINE Y1]")
    
    ObjHPosition(10) = FR_Table.Fields("[LINE X2]") + (FR_Table.Fields("[LEN H]") * 0.5) - (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(10) = FR_Table.Fields("[LINE Y2]") - FR_Table.Fields("[LEN V]")
    ObjHPosition(11) = FR_Table.Fields("[LINE X2]") + (FR_Table.Fields("[LEN H]") * 0.5) + (FR_Table.Fields("[SPACE]") * 0.5)
    ObjVPosition(11) = FR_Table.Fields("[LINE Y2]") - FR_Table.Fields("[LEN V]")
    
    REP_ABRASIVE = FR_Table.Fields("[REP]")
End Select


Select Case MARK_ANGLE
Case MARK_ANGLE_DEFAULT
        
Case MARK_ANGLE_ROTATED
            ObjHPosition(0) = FR_Table.Fields("[R TXT1 X]")
            ObjVPosition(0) = FR_Table.Fields("[R TXT1 Y]")
            ObjHPosition(1) = FR_Table.Fields("[R TXT2 X]")
            ObjVPosition(1) = FR_Table.Fields("[R TXT2 Y]")
            ObjHPosition(2) = FR_Table.Fields("[R TXT3 X]")
            ObjVPosition(2) = FR_Table.Fields("[R TXT3 Y]")
            ObjHPosition(3) = FR_Table.Fields("[R TXT4 X]")
            ObjVPosition(3) = FR_Table.Fields("[R TXT4 Y]")
            Select Case LOGO_MODE
            Case LOGO_SIDE
                ObjHPosition(4) = FR_Table.Fields("[R LOGO LX SIDE]")
                ObjVPosition(4) = FR_Table.Fields("[R LOGO LY SIDE]")
            Case LOGO_TOP
                ObjHPosition(4) = FR_Table.Fields("[R LOGO LX TOP]")
                ObjVPosition(4) = FR_Table.Fields("[R LOGO LY TOP]")
            Case LOGO_ATC
                ObjHPosition(5) = FR_Table.Fields("[R ATC  X]")
                ObjVPosition(5) = FR_Table.Fields("[R ATC  Y]")
            Case LOGO_ABRASIVE

            End Select
End Select


FR_Table.Close
FR_Database.Close

'MsgBox "SetObjProfiles Complete", vbInformation, "Laser"

'MsgBox "Complete", vbInformation, "Laser"

End Sub

Public Sub Fire_Objects(k As Integer)

Dim j As Integer
Dim i As Integer

Dim HPosition As Long
Dim VPosition As Long

For j = 0 To ObjectCount - 1
                                                                            
         '10x10 uses SEG_Y_DIST  LICA 20x20 USES LICA_XOFF,LICA_YOFF
         
        HPosition = Format((gdLocation(XLOC, k) + ObjHPosition(j) + TRAY_X_OFFSET + LICA_XOFF) * INCHES_TO_BITS, "0")
        VPosition = Format((gdLocation(YLOC, k) + ObjVPosition(j) + TRAY_Y_OFFSET + SEG_Y_DIST + LICA_YOFF) * INCHES_TO_BITS, "0")
                                                                                                                                                                                                                             
        Select Case 0
        Case 0
            
        BusyFlag = 1
        While BusyFlag = 1
            AutomationInterface.GetBusyStatus 0, BusyFlag
        Wend
        
        Select Case j
        Case 0
                
                Select Case LOGO_MODE
                Case LOGO_ABRASIVE
                
                Case Else
                        Select Case MARK_MODE
                        Case 1, 2, 3, 4
                                AutomationInterface.SetObjPos j, HPosition, VPosition
                                AutomationInterface.MarkObj j, 0
                        End Select
                End Select
        Case 1
                Select Case LOGO_MODE
                Case LOGO_ABRASIVE
                
                Case Else
                Select Case MARK_MODE
                        Case 2, 3, 4
                                AutomationInterface.SetObjPos j, HPosition, VPosition
                                AutomationInterface.MarkObj j, 0
                        End Select
                End Select
        Case 2
                Select Case LOGO_MODE
                Case LOGO_ABRASIVE
                
                Case Else
                Select Case MARK_MODE
                        Case 3, 4
                                AutomationInterface.SetObjPos j, HPosition, VPosition
                                AutomationInterface.MarkObj j, 0
                        End Select
                End Select
        Case 3
                Select Case LOGO_MODE
                Case LOGO_ABRASIVE
                
                Case Else
                        Select Case MARK_MODE
                        Case 4
                                AutomationInterface.SetObjPos j, HPosition, VPosition
                                AutomationInterface.MarkObj j, 0
                        End Select
                End Select
        Case 4
                Select Case LOGO_MODE
                Case LOGO_SIDE, LOGO_TOP
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
        Case 5
                Select Case LOGO_MODE
                Case LOGO_ATC
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        AutomationInterface.MarkObj j, 0
                End Select
          Case 6, 7, 8, 9, 10, 11
                'ABRASIVE LINE 1/2/3/4/5/6
                Select Case LOGO_MODE
                Case LOGO_ABRASIVE
                        AutomationInterface.SetObjPos j, HPosition, VPosition
                        For i = 1 To REP_ABRASIVE
                            BusyFlag2 = 1
                            While BusyFlag2 = 1
                                AutomationInterface.GetBusyStatus 0, BusyFlag2
                            Wend
                            
                            AutomationInterface.MarkObj j, 0
                            
                        Next i
                End Select
                        
        End Select
        
        End Select
        
Next j
FIRE_COUNT = FIRE_COUNT + 1
FIRE_COUNT_ID = FIRE_COUNT_ID + 1

End Sub


Public Sub Initialize_Array(Tray As Integer)

Dim X As Integer, Y As Integer, k As Integer
'
'    OFFSET TO POSITION (1,1)
'
Dim dOffSet(1) As Double

Select Case Tray
Case 9
            '
            'Lica   Case  20 X 20 Chip
            '
            For X = 0 To 9
                   For Y = 0 To 9
                        k = X + (10 * Y)
                        gdLocation(XLOC, k) = X * 0.156
                        gdLocation(YLOC, k) = Y * 0.156
                    Next Y
            Next X
            
            dOffSet(XLOC) = 0
            dOffSet(YLOC) = 0
Case 8
            '
            ' TRAY B CASE 10 X 10
            '
            '104 Tray Laser chg 12/06/2012 Array 10x10 0.181
            
            For X = 0 To 9
                   For Y = 0 To 9
                        k = X + (10 * Y)
                        gdLocation(XLOC, k) = X * 0.1815
                        gdLocation(YLOC, k) = Y * 0.181
                    Next Y
            Next X
            
            dOffSet(XLOC) = 9 * 0.5 / 2
            dOffSet(YLOC) = 9 * 0.5 / 2

Case 5, 6, 7, 10, 11
        '
        ' TRAY 301 - 412
        '
        For X = 0 To 3
               For Y = 0 To 5
                    k = X + (10 * Y)
                    Select Case X
                    Case 0 To 3
                              gdLocation(XLOC, k) = X * 0.676
                    End Select
                    Select Case Y
                    Case 0
                              gdLocation(YLOC, k) = 0
                    Case 1
                              gdLocation(YLOC, k) = 0.84
                    Case 2
                              gdLocation(YLOC, k) = 1.68
                    Case 3
                              gdLocation(YLOC, k) = 1.68 + 0.688
                    Case 4
                              gdLocation(YLOC, k) = 1.68 + 0.688 + 0.84
                    Case 5
                              gdLocation(YLOC, k) = 1.68 + 0.688 + 1.68
                    End Select
                           
                Next Y
        Next X
        
        dOffSet(XLOC) = 3 * 0.68 / 2
        dOffSet(YLOC) = (1.68 + 0.688 + 1.68) / 2

Case 2, 3, 4, 12
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
 
        dOffSet(XLOC) = 3 * 0.5 / 2
        dOffSet(YLOC) = (gdLocation(YLOC, 80) - gdLocation(YLOC, 0)) / 2

Case 1
        '
        ' TRAY 301 - 413
        '
        For X = 0 To 3
               For Y = 0 To 1
                    k = X + (10 * Y)
                    gdLocation(XLOC, k) = X * 0.68
                    gdLocation(YLOC, k) = Y * 2.375
                Next Y
        Next X
        
        dOffSet(XLOC) = 3 * 0.68 / 2
        dOffSet(YLOC) = 2.375 / 2

End Select
'
' ADD OFFSET ADJUST
'
For k = 0 To 99
    gdLocation(XLOC, k) = gdLocation(XLOC, k) - dOffSet(XLOC)
    gdLocation(YLOC, k) = -(gdLocation(YLOC, k) - dOffSet(YLOC))
Next k
   
End Sub

Public Sub Load_Job_From_File()

Screen.MousePointer = vbHourglass

Dim sFilename As String

Select Case TRAY_MARK_ANGLE
Case MARK_ANGLE_DEFAULT
        sFilename = "C:\MARKER\JOB\ATC DPSS.WLJ"
Case MARK_ANGLE_ROTATED
        sFilename = "C:\MARKER\JOB\ATC DPSS ROT.WLJ"
End Select

Dim JobIndex As Long
Beep
AutomationInterface.LoadJobFromFile sFilename, JobIndex
AutomationInterface.GetObjCount ObjectCount
 
Load_Job = 1

Screen.MousePointer = vbDefault

End Sub


Sub WorkLogNew()
 
Exit Sub
    
Set FR_Database = OpenDatabase(ATC_LASER_BD)

Dim sSQL As String
sSQL = "SELECT * FROM [TBL TRAY LOG] WHERE [TRAY_ID]=" & TRAY_ID & " AND [DATE_ID]=#" & Date & "#"

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount = 0) Then
    FR_Table.AddNew
Else
    FR_Table.Edit
End If

FR_Table.Fields("[TRAY_ID]") = TRAY_ID
FR_Table.Fields("[DATE_ID]") = Date
FR_Table.Fields("[COUNT]") = FR_Table.Fields("[COUNT]") + FIRE_COUNT
 
FR_Table.Update

FR_Table.Close
FR_Database.Close
       
FIRE_COUNT = 0
       
End Sub

Public Sub Tray_Power_Lookup(ATC_PART_ID As String)

Dim TERM_STYLE_9 As String
Dim TERM_STYLE_10 As String

TERM_STYLE_9 = Mid(ATC_PART_ID, 9, 1)
TERM_STYLE_10 = Mid(ATC_PART_ID, 10, 1)
CASE_ID = Mid(ATC_PART_ID, 4, 1)
SERIES_ID = Mid(ATC_PART_ID, 1, 3)

ValidPartNew (ATC_PART_ID)

DV_ID = Info.dDesignValue

Dim sSQL As String

Set FR_Database = OpenDatabase(ATC_LASER_BD)
Set TO_Database = OpenDatabase(ATC_LASER_BD)

sSQL = "SELECT * FROM [TBL POWER] "
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        FR_Table.Edit
        FR_Table.Fields("[PAGE]") = 0
        FR_Table.Update
        FR_Table.MoveNext
    Loop
End If
 
sSQL = "SELECT [TRAY_ID],[TBL_ID],[TERM_STYLE_9],[TERM_STYLE_10]," & _
             "[ATC PART],[COATING],[SQL_SERIES],[CASE],[VALUE],[DV MIN],[DV MAX]," & _
             "[POS 9],[POS 10 MAG],[POS 10 NON],[POS 11],[MARK PARA],[PAGE] " & _
     "FROM [TBL Power] WHERE "

sSQL = sSQL & " [ACTIVE] = Yes  ORDER BY [TRAY_ID],[ORDER]"

Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount <> 0) Then
    Do Until FR_Table.EOF
        sSQL = "SELECT [TBL_ID],[SERIES],[CASE],[VALUE],[TERM_STYLE_9],[TERM_STYLE_10],[PAGE] " & _
               "FROM [TBL Power] " & _
               "WHERE '" & CASE_ID & "'='" & FR_Table.Fields("[CASE]") & "' AND " & _
                     "'" & SERIES_ID & "' " & FR_Table.Fields("[SQL_SERIES]") & " AND " & _
                    DV_ID & ">=" & FR_Table.Fields("[DV MIN]") & " AND " & _
                    DV_ID & "<=" & FR_Table.Fields("[DV MAX]") & " AND " & _
                "'" & TERM_STYLE_9 & "' " & FR_Table.Fields("[TERM_STYLE_9]") & " AND " & _
                "'" & TERM_STYLE_10 & "' " & FR_Table.Fields("[TERM_STYLE_10]")

        Set TO_Table = TO_Database.OpenRecordset(sSQL)
        If (TO_Table.RecordCount <> 0) Then
                FR_Table.Edit
                FR_Table.Fields("[PAGE]") = 1
                TRAY_ID = FR_Table.Fields("[TRAY_ID]")
                POWER_ID = FR_Table.Fields("[TBL_ID]")
                FR_Table.Update
        End If
        FR_Table.MoveNext
    Loop
End If
TO_Database.Close
FR_Database.Close

End Sub
