Attribute VB_Name = "Common"
Option Explicit

Public Const SERVER_NY As String = "\\NY-ENG"
'039 Tape and Reel NU

Public Const SERVER_DB_NY As String = "\\NY-ENG\SPC Network\Data Base\"
Public Const SERVER_DB_JR As String = "\\Juarezdc1\Public\ATC\Data Base\"
 
Public Const DATABASE_MODE_REM_NY As Integer = 0        ' NY MODE NY NOTE
                                                        ' DEFAULT IS ZERO REMOTE NY
Public Const DATABASE_MODE_REM_JR As Integer = 3        ' NY MODE JUAREZ DATABASE VIEW
Public Const DATABASE_MODE_REM_JUAREZ As Integer = 4    ' JUAREZ EXCLUSIVE MODE
Public Const DATABASE_MODE_FIL As Integer = 2           ' CONFIGURATION FILE
Public Const DATABASE_MODE_LCL As Integer = 1           ' LOCAL
Public DataBase_MODE As Integer

Public Const MASTER_PASSWORD As String = "JR2013"

Public COUNT_TIME As Long
'039 Tape and Reel
'084 ESI 3340
'048 AutoSort

Public LOCATION_ID As String          ' NY/JR
Public DEPT_ID As String
Public OP_ID As Long
Public MACHINE_ID As Long

Public MACHINE_TYPE As Long         '[7:LASER][11:TAPE&REEL][1:ESI 3340][2:PAL18A]
Public MACHINE_NUMBER As String     'PART
Public MACHINE_DESCRIPTION As String


Public WS_ID As Long        'WORK SHEET ID
Public WO_ID As String      'WORK ORDER

Public CODE_ID As Long
Public DF_ID As Long
Public DEFECT_ID As Long

Public DATE_ID As String
Public DATE_START_ID As String
Public DATE_END_ID As String

'039 Tape and Reel
'092 Visual Inspection
'110 Visual Inspection
'088 WR OEE
'100 Close Out
Type Info_Type
   
            iSampleSize As Integer
            sCalDate As String * 10
            iSerialNumber As Integer
            
            iLambdaQty As Integer       ' Number of Lambdas to Test Minimum 1
            iLambdaNum(10) As Integer   ' Lambda Number
            dLambdaESR(10) As Double    ' ESR Limit
            iUniqueNumber  As Integer
            iMachineNumber As Integer
            sConfiguration As String
            sProduct As String
            sDate As String
            sTime  As String
            dTime As Double
            sFile As String
            sPassword As String
            iPresetCount As Integer
            iCapSampleSize As Integer
            iIRSampleSize As Integer
            iDWVSampleSize As Integer
            iRealSampSize As Integer
            lBatch_No As Long
            
            iRotaryNo As Integer
            
            sOperator As String
            sOrderNumber As String
            sLotNumber As String
            sLERnumber As String
            sNote As String
            sRevDate As String
            sProdTest As String
            
            
            sATCPart As String
            sSERIES As String
            sSeriesCase As String
            sCASE As String
            dDesignValue As Double
            sDesignValue As String
            sDesignTolerance As String
            
            
            lLotQuantity As Long
            lQuantity As Long
            
            iTableType As Integer
            sTableTypeName As String
            sTol As String
            
            sTolerance As String
            iStopCounter As Integer       'How many times stopped during normal operation
            
            dDFLimit As Double
            
            iDF_TEST_MODE As Integer
            dDF_Low As Double
            dDF_High As Double
            
            sTestFreq As String
            dStandRef_C As Double
            dStandRef_DF As Double
            
            dOpenComp_C As Double
            dShortComp_R As Double
End Type
Public Info As Info_Type


Public Const FREAD As Integer = 0
Public Const FWRITE As Integer = 1
Public Const FDATA As Integer = 2

' ==========================================================
' = Get Memory Information                                =
' ==========================================================
Public Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)

Private Type INT64
   LoPart As Long
   HiPart As Long
End Type

Public Type MEMORYSTATUSEX
   dwLength As Long
   dwMemoryLoad As Long
   ulTotalPhys As INT64
   ulAvailPhys As INT64
   ulTotalPageFile As INT64
   ulAvailPageFile As INT64
   ulTotalVirtual As INT64
   ulAvailVirtual As INT64
   ulAvailExtendedVirtual As INT64
End Type


' ==========================================================
' = Get Windows Information                                =
' ==========================================================

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2


Public strComputerName As String
Public IP_ADDRESS As String
Public COMP_ID As Long

Public Const MAX_COMPUTERNAME_LENGTH = 31
Public Const MAX_PATH = 260
Public Const UNLEN = 256

Public Declare Function GetComputerName Lib "kernel32" _
   Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
      
Public Declare Function GetUserName Lib "advapi32.dll" _
   Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long
'
'
'
Public Sub Get_User()

Dim lngReturn As Long
Dim lngLen As Long
Dim strString As String
Dim strUserName As String

On Error GoTo DB_Error

strString = String(CInt(MAX_COMPUTERNAME_LENGTH + 1), Chr(0))
lngLen = MAX_COMPUTERNAME_LENGTH
lngReturn = GetComputerName(strString, lngLen)

strComputerName = Left(strString, lngLen)

strString = String(UNLEN + 1, Chr(0))
lngLen = UNLEN
lngReturn = GetUserName(strString, lngLen)
strUserName = Left(strString, lngLen - 1)
 
Exit Sub
DB_Error:
     
Exit Sub
 
End Sub
'
'
'
Public Sub ConfigComputer_DB(iModeStatus As Integer)
 
On Error GoTo Error_Config
 
Const COMPUTER_DB_NY As String = "\\NY-ENG\SPC Network\Data Base\COMPUTER CONFIG.mdb"
Const COMPUTER_DB_JR As String = "\\Juarezdc1\Public\ATC\Data Base\COMPUTER CONFIG JR.mdb"
 
Select Case LOCATION_ID
Case "NY"
            Set FR_Database = OpenDatabase(COMPUTER_DB_NY)
Case "JR"
            Set FR_Database = OpenDatabase(COMPUTER_DB_JR)
End Select
       
Dim sSQL As String

sSQL = "SELECT * FROM [CONFIG] WHERE [COMPUTER] ='" & strComputerName & "'"
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount = 0) Then
        FR_Table.AddNew
        FR_Table.Fields("[COMPUTER]") = strComputerName
Else
        FR_Table.Edit
End If

FR_Table.Fields("[SCREEN]") = Screen.Width / Screen.TwipsPerPixelX & " by " & Screen.Height / Screen.TwipsPerPixelY
FR_Table.Fields("[DATE START]") = Format(Date, "MM/dd/yyyy ") & Format(Time, "hh:mm am/pm")
FR_Table.Fields("[IP]") = IP_ADDRESS
    
FR_Table.Fields("[OP SYS]") = Mid(GetWindowsVersion, 1, 20)
  
  
Dim udtMemStatEx As MEMORYSTATUSEX
udtMemStatEx.dwLength = Len(udtMemStatEx)
GlobalMemoryStatusEx udtMemStatEx

FR_Table.Fields("[RAM]") = NumberInKB(CLargeInt(udtMemStatEx.ulTotalPhys.LoPart, udtMemStatEx.ulTotalPhys.HiPart))
     

COMP_ID = FR_Table.Fields("[ID]")
                  
FR_Table.Update
                                                            
sSQL = "SELECT * FROM [TBL ATC DWG] WHERE [ATC DWG] ='" & TBL_ATC_DWG & "'"
Set FR_Table = FR_Database.OpenRecordset(sSQL)
          
If (FR_Table.RecordCount = 0) Then
        FR_Table.AddNew
        FR_Table.Fields("[ATC DWG]") = TBL_ATC_DWG
        FR_Table.Fields("[TITLE]") = TBL_NAME
Else
        FR_Table.Edit
End If

Dim DWG_ID As Long
DWG_ID = FR_Table.Fields("[DWG_ID]")

FR_Table.Fields("[DATE_ID]") = ATC_VERSION
FR_Table.Update

sSQL = "SELECT * FROM [TBL DWG USER] " & _
       "WHERE [DWG_ID] =" & DWG_ID & " AND [CP_ID]=" & COMP_ID
       
Set FR_Table = FR_Database.OpenRecordset(sSQL)
If (FR_Table.RecordCount = 0) Then
        FR_Table.AddNew
        FR_Table.Fields("[DWG_ID]") = DWG_ID
        FR_Table.Fields("[CP_ID]") = COMP_ID
Else
        FR_Table.Edit
End If

FR_Table.Fields("[DATE_ID]") = Format(Date, "MM/dd/yyyy ") & Format(Time, "hh:mm am/pm")

Select Case iModeStatus
Case 0
           FR_Table.Fields("[STATUS]") = "Open"
Case 1
           FR_Table.Fields("[STATUS]") = "Closed"
Case 2
           FR_Table.Fields("[STATUS]") = "Exit"
End Select

FR_Table.Fields("[REV DATE]") = ATC_VERSION
FR_Table.Update

FR_Database.Close

Exit Sub
Error_Config:
     
Resume Next

End Sub

Public Function GetIPAddress() As String

    Dim buf(0 To 511) As Byte
    Dim BufSize As Long
    Dim lngResult As Long
    Dim intAddress As Integer
    Dim intOctet As Integer
    Dim strAddress(0 To 10) As String
    Dim strTempAddress As String
    
    BufSize = 512
    
    lngResult = GetIpAddrTable_API(buf(0), BufSize, 1)
    
    'FUNCTION RETURNED AN ERROR
    If lngResult <> 0 Then
        GetIPAddress = "Unknown"
        Exit Function
    End If
    'MORE THAN 256 IP ADDRESSES FOUND
    If buf(1) <> 0 Then
        GetIPAddress = "Unknown"
        Exit Function
    End If
    'NO IP ADDRESSES FOUND OR TOO MANY
    If buf(0) = 0 Or buf(0) > UBound(strAddress) Then
        GetIPAddress = "Unknown"
        Exit Function
    End If

    'EXTRACT IP ADDRESS FROM BUF
    Dim strIPAddress As String
    For intAddress = 0 To buf(0) - 1
        strTempAddress = ""
        For intOctet = 0 To 3
            strTempAddress = strTempAddress & IIf(intOctet > 0, ".", "") & buf(4 + intAddress * 24 + intOctet)
        Next intOctet
        strAddress(intAddress) = strTempAddress
        Select Case strTempAddress
        Case "", "127.0.0.1"
        Case Else
                strIPAddress = strTempAddress
        End Select
    Next
    If strAddress(0) <> "" Then
        'GetIPAddress = strAddress(0)
        GetIPAddress = strIPAddress
    Else
        GetIPAddress = "Unknown"
    End If

End Function
'
'   vbPRORLandscape     FONT SIZE 10
'   vbPRORPortrait
'
Public Sub PrintFile(sFilename As String, iOrientation As Integer, iFontSize As Integer)

Dim strFont As String, sngSize As Single

Dim iFilenum As Integer
Dim stemp As String
 
' SAVE CURRENT PRINTER SETTINGS
strFont = Printer.Font
sngSize = Printer.FontSize
 
Printer.Orientation = iOrientation
Printer.Font = "Courier New"
Printer.FontSize = iFontSize

iFilenum = FreeFile
Open sFilename For Input Shared As iFilenum
Do Until EOF(iFilenum)
        Line Input #iFilenum, stemp
        Printer.Print stemp
Loop
Close iFilenum

Printer.EndDoc
 
' RESET PRINTER SETTINGS
Printer.Font = strFont
Printer.FontSize = sngSize

End Sub
'
'   ONLY ACCURATE TO APPROX. 0.1 SECOND INTERVALS (RESOLUTION)
'     The number of seconds elapsed since midnight.
'     24 * 60 * 60 = 86400
'
Sub Pause(dDelayTime As Double)

Dim dCurrentTime As Double

Dim dStartTime As Double
dStartTime = Timer
Do
    DoEvents
    dCurrentTime = Timer
    If (dCurrentTime >= dStartTime) Then
        If (dCurrentTime - dStartTime >= dDelayTime) Then
            Exit Do
        End If
    Else
       ' Just Past midnight
       If (dCurrentTime + (86400 - dStartTime) >= dDelayTime) Then
            Exit Do
       End If
    End If
Loop

End Sub


' Returns the version of Windows that the user is running
Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s Win 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"

                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Win NT 3.5"
                    Case 4
                        GetWindowsVersion = "Win NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Win 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Win Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Win Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Win 7/Server 2008 R2"
                        End Select
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Win Unknown"
    End If
End Function


Public Function NumberInKB(ByVal vNumber As Currency) As String
   Dim strReturn As String

   Select Case vNumber
      Case Is < 1024 ^ 1
         strReturn = CStr(vNumber) & " bytes"

      Case Is < 1024 ^ 2
         strReturn = CStr(Round(vNumber / 1024, 1)) & " KB"

      Case Is < 1024 ^ 3
         strReturn = CStr(Round(vNumber / 1024 ^ 2, 2)) & " MB"

      Case Is < 1024 ^ 4
         strReturn = CStr(Round(vNumber / 1024 ^ 3, 2)) & " GB"
   End Select

   NumberInKB = strReturn

End Function

'This function converts the LARGE_INTEGER data type to a double
Public Function CLargeInt(Lo As Long, Hi As Long) As Double
   Dim dblLo As Double
   Dim dblHi As Double

   If Lo < 0 Then
      dblLo = 2 ^ 32 + Lo
   Else
      dblLo = Lo
   End If

   If Hi < 0 Then
      dblHi = 2 ^ 32 + Hi
   Else
      dblHi = Hi
   End If

   CLargeInt = dblLo + dblHi * 2 ^ 32

End Function


Function InfoVal(stemp As String) As Double

InfoVal = Val(Mid(stemp, InStr(1, stemp, "=") + 1, Len(stemp) - InStr(1, stemp, "=")))

End Function
'
Function InfoStr(stemp As String) As String

InfoStr = Mid(stemp, InStr(1, stemp, "=") + 1, Len(stemp) - InStr(1, stemp, "="))

End Function


Function InfoValStr(stemp As String, sChar As String) As Double
 
InfoValStr = Val(Mid(stemp, InStr(1, stemp, sChar) + 1, Len(stemp) - InStr(1, stemp, sChar)))

End Function
'
Function InfoStrChar(stemp As String, sChar As String) As String

InfoStrChar = Mid(stemp, InStr(1, stemp, sChar) + 1, Len(stemp) - InStr(1, stemp, sChar))

InfoStrChar = LTrim(InfoStrChar)

End Function


'
'
'   Lot Decode  DWG 108-1192
'               MANUFACTURING LOT NUMBER
'               IDENTIFICATION SYSTEM PROCEDURE
'
'   RETURNS ATCPART SERIES,CASE,VALUE AS FORMATTED STRING
'
'   chg 10/28/2009 P K22 800 SERIES
'
Function LotDecode(sLotNumber As String) As String

    'Test Format A11AXXXAAA

   Dim sSERIES As String
   LotDecode = ""
  
  '1 SERIES (PRODUCT LINE)
   Select Case Mid$(sLotNumber, 1, 1)
   Case "A", "E", "N"
                sSERIES = "100"
   Case "C", "D"
                sSERIES = "700"
   Case "F"
                sSERIES = "200"
   Case "K"
                sSERIES = "600"
   Case "P"
                sSERIES = "800"
   Case "R"
                sSERIES = "180"
   Case "X"
                sSERIES = "200"
                If (Mid$(sLotNumber, 10, 1) = "C") Then
                          sSERIES = "900"
                End If
   Case "Z"
                sSERIES = "710"
                If (Mid$(sLotNumber, 10, 1) = "R") Then
                          sSERIES = "180"
                End If
   Case Else
                Exit Function
   End Select
   
    '10 ATC CASE CODE
    Dim sCASE As String
    sCASE = Mid$(sLotNumber, 10, 1)
    Select Case Mid$(sLotNumber, 10, 1)
    Case "A"
            sCASE = "A"
    Case "B"
            sCASE = "B"
    Case "C"
            sCASE = "C"
    Case "E"
            sCASE = "E"
    Case "T"
            sCASE = "B"
    Case "R"
            sCASE = "R"
    Case "U"
            sCASE = "A"
    Case Else
    End Select
     
     
    Select Case Mid$(sLotNumber, 12, 1) 'Ceramic_Material
    Case "E"
            sSERIES = "900"
    Case "F"
            sSERIES = "700"
    End Select
   '567 CAPACITANCE CODE
  
   LotDecode = sSERIES & sCASE & Mid$(sLotNumber, 5, 3)
   
   Info.sSeriesCase = sSERIES & sCASE

   '---  CONVERT TO A DOUBLE PF VALUE
   If (Mid$(LotDecode, 6, 1) = "R") Then
      Info.dDesignValue = Val(Mid$(LotDecode, 5, 1) & "." & Mid$(LotDecode, 7, 1))
   Else
      Info.dDesignValue = Val(Mid$(LotDecode, 5, 2) & "E" & Mid$(LotDecode, 7, 1))
   End If

   Info.sDesignValue = Str$(Info.dDesignValue)


End Function
'
'   Valid Part Sort checks for valid part format
'   Series Format       NNN
'   CASE SIZE           L
'   Cap Value Format    NNN NRN
'   No Tolerance
'
Function ValidPart(sATCPart As String) As Boolean
      
   Info.sATCPart = sATCPart
   Info.sSeriesCase = Mid$(sATCPart, 1, 4)  '-- 100E
   Info.sSERIES = Mid$(sATCPart, 1, 3)
   Info.sCASE = Mid$(sATCPart, 4, 1)
            
   If (Len(sATCPart) >= 8) Then
        Info.sTol = Mid$(sATCPart, 8, 1)
   End If
   '----- TEST FOR VALID SERIES DESIGN VALUE FORMAT
   Select Case Mid$(sATCPart, 1, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPart = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 2, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPart = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 3, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPart = False
                Exit Function
   End Select
   
'---- TEST FOR VALID SERIES
   Select Case Info.sSeriesCase
   Case "100A", "100B", "100C", "100E"
   Case "110A"
   Case "175B", "180R"
   Case "200A", "200B"
   Case "500S"
   Case "600A", "600B", "600C", "600E", "600F", "600L", "600R", "600S"
   Case "650S", "650F", "650L"
   Case "700A", "700B", "700E"
   Case "710B", "710A", "710C", "710E"
   Case "800C", "800E"
   Case "900B", "900C"
   
   Case Else
              '   MsgBox "Series Not Valid", vbInformation, "ATC Part Number"
              '  ValidPart = False
              '  Exit Function
   End Select
         
   Select Case Mid$(sATCPart, 4, 1)
   Case "A", "B", "C", "E", "R", "T", "S", "L", "F"
   Case Else
               ' ValidATC = False
               ' Exit Function
   End Select
      
   Select Case Mid$(sATCPart, 4, 1)
   Case "A" To "Z"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPart = False
                Exit Function
   End Select

   '----- TEST FOR VALID DESIGN VALUE FORMAT

   Select Case Mid$(sATCPart, 5, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
                ValidPart = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 6, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "R"
   Case Else
                MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
                ValidPart = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 7, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
                ValidPart = False
                Exit Function
   End Select

   '---  CONVERT TO A DOUBLE PF VALUE
   If (Mid$(sATCPart, 6, 1) = "R") Then
      Info.dDesignValue = Val(Mid$(sATCPart, 5, 1) & "." & Mid$(sATCPart, 7, 1))
   Else
      Info.dDesignValue = Val(Mid$(sATCPart, 5, 2) & "E" & Mid$(sATCPart, 7, 1))
   End If

   Info.sDesignValue = Str$(Info.dDesignValue)
   
   ValidPart = True

End Function


'
'   Valid Part Format No returns except Valid
'
Function ValidPartX(sATCPart As String) As Boolean

   
   '----- TEST FOR VALID SERIES DESIGN VALUE FORMAT
   Select Case Mid$(sATCPart, 1, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPartX = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 2, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPartX = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 3, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPartX = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 4, 1)
   Case "A" To "Z"
   Case Else
                MsgBox "Cap Series Format Not Valid", vbInformation, "ATC Part Number"
                ValidPartX = False
                Exit Function
   End Select

   '----- TEST FOR VALID DESIGN VALUE FORMAT

   Select Case Mid$(sATCPart, 5, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
                ValidPartX = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 6, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "R"
   Case Else
                MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
                ValidPartX = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 7, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                MsgBox "Cap Format Not Valid", vbInformation, "ATC Part Number"
                ValidPartX = False
                Exit Function
   End Select
   
   ValidPartX = True

End Function
'
'   Format Numbers 6 Significant Figures
'
Function FormatStr(dNumber As Double) As String

If Mid$(Info.sATCPart, 6, 1) = "R" Then
    If (Mid$(Info.sATCPart, 5, 1) = "0") Then
            FormatStr = Format$(dNumber, "0.000000")
    Else
            FormatStr = Format$(dNumber, "0.00000")
    End If
Else
    Select Case Mid$(Info.sATCPart, 7, 1)
    Case "0"
            FormatStr = Format$(dNumber, "#0.0000")
    Case "1"
            FormatStr = Format$(dNumber, "##0.000")
    Case "2"
            FormatStr = Format$(dNumber, "###0.00")
    Case "3"
            FormatStr = Format$(dNumber, "#####0.0")
    Case Else
            FormatStr = Format$(dNumber, "######0")
    End Select
End If

If (dNumber = 2E+32) Then
    FormatStr = "UNBAL"
End If

End Function
'
'   Valid Part Sort checks for valid part format
'   Series , Cap Value Format , No Tolerance
'
Function ValidPartShort(sATCPart As String) As Boolean
      
   Dim sBuff As String
      
   Info.sATCPart = sATCPart
   Info.sSeriesCase = Mid$(sATCPart, 1, 4)  '-- 100E
   
   
   '----- TEST FOR VALID SERIES DESIGN VALUE FORMAT
   Select Case Mid$(sATCPart, 1, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ValidPartShort = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 2, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ValidPartShort = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 3, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ValidPartShort = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 4, 1)
   Case "A", "B", "C", "E", "R"
   Case Else
                ValidPartShort = False
                Exit Function
   End Select

   '----- TEST FOR VALID DESIGN VALUE FORMAT

   Select Case Mid$(sATCPart, 5, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ValidPartShort = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 6, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "R"
   Case Else
                ValidPartShort = False
                Exit Function
   End Select
   Select Case Mid$(sATCPart, 7, 1)
   Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
   Case Else
                ValidPartShort = False
                Exit Function
   End Select

   '---  CONVERT TO A DOUBLE PF VALUE
   If (Mid$(sATCPart, 6, 1) = "R") Then
      Info.dDesignValue = Val(Mid$(sATCPart, 5, 1) & "." & Mid$(sATCPart, 7, 1))
   Else
      Info.dDesignValue = Val(Mid$(sATCPart, 5, 2) & "E" & Mid$(sATCPart, 7, 1))
   End If

   Info.sDesignValue = Str$(Info.dDesignValue)
   
   ValidPartShort = True

End Function


Public Sub AddFields()

On Error GoTo Error_AddFields

Dim dbs As Database
'Set dbs = OpenDatabase(SERVER_DB_JR & "OEE SPM JR MASTER.MDB")

' ADD FIELDS INTEGER,BYTE,LONG,SINGLE,TEXT(4)

dbs.Execute "ALTER TABLE [Defect List] ADD COLUMN [TR] TEXT(2);"
dbs.Execute "ALTER TABLE [Defect List] ADD COLUMN [LS] TEXT(2);"
dbs.Execute "ALTER TABLE [Defect List] ADD COLUMN [ET] TEXT(2);"

dbs.Execute "ALTER TABLE [INFO] ADD COLUMN [OPEN STATUS] TEXT(4);"
dbs.Execute "ALTER TABLE [INFO] ADD COLUMN [SHORT STATUS] TEXT(4);"
dbs.Execute "ALTER TABLE [INFO] ADD COLUMN [STAND STATUS] TEXT(4);"
dbs.Execute "ALTER TABLE [INFO] ADD COLUMN [OPEN COMP] SINGLE;"
dbs.Execute "ALTER TABLE [INFO] ADD COLUMN [SHORT COMP] SINGLE;"
dbs.Execute "ALTER TABLE [INFO] ADD COLUMN [STANDARD REF DF] SINGLE;"
dbs.Execute "ALTER TABLE [INFO] ADD COLUMN [STANDARD REF CAP] SINGLE;"

dbs.Close

Exit Sub

Error_AddFields:

Resume Next

End Sub

Public Function Lot_Number_Decode(sLotNumber As String) As String

    Dim sSERIES As String

    Select Case Mid$(sLotNumber, 1, 1)
    Case "A", "E", "N"
                 sSERIES = "100"
    Case "C", "D"
                 sSERIES = "700"
    Case "F"
                 sSERIES = "200"
    Case "K"
                 sSERIES = "600"
    Case "P"
                 sSERIES = "800"
    Case "R"
                 sSERIES = "180"
    Case "X"
                 sSERIES = "200"
                 If (Mid$(sLotNumber, 10, 1) = "C") Then
                           sSERIES = "900"
                 End If
    Case "Z"
                 sSERIES = "710"
                 If (Mid$(sLotNumber, 10, 1) = "R") Then
                           sSERIES = "180"
                 End If
    Case Else
                 Lot_Number_Decode = " "
                 Exit Function
    End Select
   
    '10 ATC CASE CODE
    Dim sCASE As String
    sCASE = Mid$(sLotNumber, 10, 1)
    Select Case Mid$(sLotNumber, 10, 1)
    Case "A"
            sCASE = "A"
    Case "B"
            sCASE = "B"
    Case "C"
            sCASE = "C"
    Case "E"
            sCASE = "E"
    Case "T"
            sCASE = "B"
    Case "R"
            sCASE = "R"
    Case "U"
            sCASE = "A"
    Case Else
            Lot_Number_Decode = " "
    End Select
     
  '567 CAPACITANCE CODE
  
     Lot_Number_Decode = sSERIES & sCASE & Mid$(sLotNumber, 5, 3)
               
End Function


Function InfoHeaderStr(stemp As String) As String

If (stemp = "") Then
    Exit Function
End If

InfoHeaderStr = Mid(stemp, 1, InStr(1, stemp, "=") - 1)

End Function

