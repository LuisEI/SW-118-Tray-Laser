Attribute VB_Name = "Laser_101"

Option Explicit
'
'
'
Public Sub InputOutput(frmForm As Form)

'===================================================================================

'chg 01/25/2013 InputOutput IO Definition

'       DIGITAL I/O MODULE RACK
'       Control Laser Corporation
'
'       Input Defintion
'===================================================================================
                                                       
frmForm.lblInput(1).ToolTipText = "Position 1     TSTBIT 1"     'Position 1     TSTBIT 1
frmForm.lblInput(2).ToolTipText = "Position 2     TSTBIT 2"     'Position 2     TSTBIT 2
frmForm.lblInput(3).ToolTipText = "Position 3     TSTBIT 3"     'Position 3     TSTBIT 3
frmForm.lblInput(4).ToolTipText = "Position 4     TSTBIT 4"     'Position 4     TSTBIT 4
frmForm.lblInput(5).ToolTipText = "Position 5     TSTBIT 5"     'Position 5     TSTBIT 5
frmForm.lblInput(6).ToolTipText = "Position 6     TSTBIT 6"     'Position 6     TSTBIT 6
frmForm.lblInput(7).ToolTipText = "Position 7     TSTBIT 7"     'Position 7     TSTBIT 7
frmForm.lblInput(8).ToolTipText = "Position 8     TSTBIT 8"     'Position 7     TSTBIT 8

frmForm.lblInput(1).Caption = "Dial 1 (A) Part Present"             'Position 1     TSTBIT 1
frmForm.lblInput(2).Caption = "Dial 1 (A) Error 1"                  'Position 2     TSTBIT 2
frmForm.lblInput(3).Caption = "Dial 1 (A) Error 2"                  'Position 3     TSTBIT 3
frmForm.lblInput(4).Caption = "Dial 2 (B) Part Present"             'Position 4     TSTBIT 4
frmForm.lblInput(5).Caption = "Dial 2 (B) Error 1"                  'Position 5     TSTBIT 5
frmForm.lblInput(6).Caption = "Dial 2 (B) Error 2"                  'Position 6     TSTBIT 6
frmForm.lblInput(7).Caption = "Interlocks"                          'Position 7     TSTBIT 7
frmForm.lblInput(8).Caption = "Dial 1(A)/2(B)"                      'Position 8     TSTBIT 8

'===================================================================================
'       Output Defintion
'===================================================================================
                                                                
frmForm.cmdOut(1).Caption = "Main Air"                          'Position 1     OUTBIT 1,1/0
frmForm.cmdOut(2).Caption = "Air 1(A)/2(B)"                     'Position 2     OUTBIT 2,1/0
frmForm.cmdOut(3).Caption = "Bowl 1 (A)"                        'Position 3     OUTBIT 3,1/0
frmForm.cmdOut(4).Caption = "Bowl 2 (B)"                        'Position 4     OUTBIT 4,1/0
frmForm.cmdOut(5).Caption = "Spare"                             'Position 5     OUTBIT 5,1/0
frmForm.cmdOut(6).Caption = "Spare"                             'Position 6     OUTBIT 6,1/0
frmForm.cmdOut(7).Caption = "Spare"                             'Position 7     OUTBIT 7,1/0
frmForm.cmdOut(8).Caption = "Spare"                             'Position 8     OUTBIT 8,1/0

frmForm.cmdOut(1).ToolTipText = "Position 1    OUTBIT 1,1/0"
frmForm.cmdOut(2).ToolTipText = "Position 2    OUTBIT 2,1/0"
frmForm.cmdOut(3).ToolTipText = "Position 3    OUTBIT 3,1/0"
frmForm.cmdOut(4).ToolTipText = "Position 4    OUTBIT 4,1/0"
frmForm.cmdOut(5).ToolTipText = "Position 5    OUTBIT 5,1/0"
frmForm.cmdOut(6).ToolTipText = "Position 6    OUTBIT 6,1/0"
frmForm.cmdOut(7).ToolTipText = "Position 7    OUTBIT 7,1/0"
frmForm.cmdOut(8).ToolTipText = "Position 8    OUTBIT 8,1/0"

End Sub


Public Sub ShowIO(frmForm As Form)

   Dim i As Integer
   
   For i = 1 To 8
        frmForm.lblInp(i).Caption = iInput(i)
        Select Case iInput(i)
        Case 1
                frmForm.lblInp(i).BackColor = &HC0FFC0
        Case 0
                frmForm.lblInp(i).BackColor = &HC0C0FF
        End Select
        
    Next i
 
    If (frmForm.lblInp(1).FontBold = False) Then
        For i = 1 To 8
            frmForm.lblInp(i).FontBold = True
        Next i
    Else
        For i = 1 To 8
            frmForm.lblInp(i).FontBold = False
        Next i
    End If
    
 
    Dim InputHi As Integer
    Select Case CONFIGURATION_ID
    Case 0
            InputHi = 1
    Case 1
            InputHi = 0
    End Select
    
    If (iInput(1) = InputHi) Then
        frmForm.lblInpV(1).Caption = "PP"
    Else
        frmForm.lblInpV(1).Caption = "NP"
    End If
    If (iInput(2) = InputHi) Then
        frmForm.lblInpV(2).Caption = "Error"
    Else
        frmForm.lblInpV(2).Caption = "Clear"
    End If
    If (iInput(3) = InputHi) Then
        frmForm.lblInpV(3).Caption = "Error"
    Else
        frmForm.lblInpV(3).Caption = "Clear"
    End If
    
    If (iInput(4) = InputHi) Then
        frmForm.lblInpV(4).Caption = "PP"
    Else
        frmForm.lblInpV(4).Caption = "NP"
    End If
    If (iInput(5) = InputHi) Then
        frmForm.lblInpV(5).Caption = "Error"
    Else
        frmForm.lblInpV(5).Caption = "Clear"
    End If
    If (iInput(6) = InputHi) Then
        frmForm.lblInpV(6).Caption = "Error"
    Else
        frmForm.lblInpV(6).Caption = "Clear"
    End If
    If (iInput(7) = InputHi) Then
        frmForm.lblInpV(7).Caption = "Closed"
    Else
        frmForm.lblInpV(7).Caption = "Open"
    End If
    If (iInput(8) = InputHi) Then
        frmForm.lblInpV(8).Caption = "2(B)"
    Else
        frmForm.lblInpV(8).Caption = "1(A)"
    End If

End Sub

Public Sub SetOutputs(frmForm As Form, iMode As Integer)
 
Select Case iMode
Case 0
 
Case 2
        'INITIALIZE
        frmForm.lblOut(1).Caption = "0"
        frmForm.lblOut(2).Caption = "0"
        frmForm.lblOut(3).Caption = "0"
        frmForm.lblOut(4).Caption = "0"
        frmForm.lblOut(5).Caption = "0"
        frmForm.lblOut(6).Caption = "0"
        frmForm.lblOut(7).Caption = "0"
                                            
End Select
 
End Sub
