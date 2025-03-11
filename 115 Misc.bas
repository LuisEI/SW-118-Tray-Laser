Attribute VB_Name = "LibraryFunctions"
Option Explicit
'
'
Public Function Covert_to_EIA(sATCPart As String) As String
       
'----- TEST FOR VALID SERIES DESIGN VALUE FORMAT
Dim ValidPart As Boolean

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

Dim dDesignValue As Double

If (Mid$(sATCPart, 6, 1) = "R") Then
   'UNDER 10 PF     X.X FORMAT
   dDesignValue = Val(Mid$(sATCPart, 5, 1) & "." & Mid$(sATCPart, 7, 1))
Else
   dDesignValue = Val(Mid$(sATCPart, 5, 2) & "E" & Mid$(sATCPart, 7, 1))
End If


ValidPart = True
 
Dim EIA_SIGN_FIG As String
'---  CONVERT TO A DOUBLE PF VALUE
Dim SIGN_FIG As String
Dim EIA_MULT As Integer

Select Case dDesignValue
Case Is < 1
        '0.X x 10E0  FORMAT NRN
        '0.X    MULT -1
        EIA_SIGN_FIG = Format(dDesignValue * 10, "0.0")
        EIA_MULT = -1
Case Is < 10
        'X.X x 10E0  FORMAT NRN
        'X.X    MULT 0
        EIA_SIGN_FIG = Format(dDesignValue, "0.0")
        EIA_MULT = 0
Case Else
        ' >10
        '
        ' X.X MULT = M-1
        EIA_SIGN_FIG = Mid$(sATCPart, 5, 1) & "." & Mid$(sATCPart, 6, 1)
        EIA_MULT = Mid$(sATCPart, 7, 1) + 1
End Select
    
Dim sSQL As String

Set FR_Database = OpenDatabase(DB_DPSS_LASER)

sSQL = "SELECT * FROM [1STCHAR] WHERE [SIGN FIG] = '" & EIA_SIGN_FIG & "'"

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    Covert_to_EIA = FR_Table.Fields("ALPHA CHAR")
Else
    Covert_to_EIA = "Not Valid"
    FR_Table.Close
    FR_Database.Close
    Exit Function
End If
sSQL = "SELECT * FROM [2ND CHAR] WHERE [DEC MULT] = " & EIA_MULT

Set FR_Table = FR_Database.OpenRecordset(sSQL)

If (FR_Table.RecordCount <> 0) Then
    Covert_to_EIA = Covert_to_EIA & FR_Table.Fields("NUM CHAR")
Else
    Covert_to_EIA = "Not Valid"
    FR_Table.Close
    FR_Database.Close
    Exit Function
End If

FR_Table.Close
FR_Database.Close
    
End Function
