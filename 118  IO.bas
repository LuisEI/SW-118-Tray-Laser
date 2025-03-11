Attribute VB_Name = "IO_Module"
Option Explicit

Public BOARD_IO As Integer
'
'   OUTPUT PORT CONSTANTS
'
Public Const OUTPUT_1 As Integer = 0
Public Const OUTPUT_2 As Integer = 1
Public Const OUTPUT_3 As Integer = 2
Public Const OUTPUT_4 As Integer = 3
Public Const OUTPUT_5 As Integer = 4
Public Const OUTPUT_6 As Integer = 5
Public Const OUTPUT_7 As Integer = 6
Public Const OUTPUT_8 As Integer = 7


Public Const OUTPUT_1A As Integer = 8
Public Const OUTPUT_2A As Integer = 9
Public Const OUTPUT_3A As Integer = 10
Public Const OUTPUT_4A As Integer = 11
Public Const OUTPUT_5A As Integer = 12
Public Const OUTPUT_6A As Integer = 13
Public Const OUTPUT_7A As Integer = 14
Public Const OUTPUT_8A As Integer = 15


Public Const OUTPUT_1B As Integer = 16
Public Const OUTPUT_2B As Integer = 17
Public Const OUTPUT_3B As Integer = 18
Public Const OUTPUT_4B As Integer = 19
Public Const OUTPUT_5B As Integer = 20
Public Const OUTPUT_6B As Integer = 21
Public Const OUTPUT_7B As Integer = 22
Public Const OUTPUT_8B As Integer = 23

Public OUTPUT_Status(30) As Integer

Public Const Init_IO As Integer = 24


'------------------------------------------------------------
' STATUS READ OPTIONS AND INPUTS
'------------------------------------------------------------
Public Const INPUT_1A As Integer = 0
Public Const INPUT_2A As Integer = 1
Public Const INPUT_3A As Integer = 2
Public Const INPUT_4A As Integer = 3
Public Const INPUT_5A As Integer = 4
Public Const INPUT_6A As Integer = 5
Public Const INPUT_7A As Integer = 6
Public Const INPUT_8A As Integer = 7

Public Const INPUT_1B As Integer = 8
Public Const INPUT_2B As Integer = 9
Public Const INPUT_3B As Integer = 10
Public Const INPUT_4B As Integer = 11
Public Const INPUT_5B As Integer = 12
Public Const INPUT_6B As Integer = 13
Public Const INPUT_7B As Integer = 14
Public Const INPUT_8B As Integer = 15


Public Const INPUT_1C As Integer = 16
Public Const INPUT_2C As Integer = 17
Public Const INPUT_3C As Integer = 18
Public Const INPUT_4C As Integer = 19
Public Const INPUT_5C As Integer = 20
Public Const INPUT_6C As Integer = 21
Public Const INPUT_7C As Integer = 22
Public Const INPUT_8C As Integer = 23
'------------------------------------------------------------
 

Sub Init_PCI_IO()

   ' declare revision level of Universal Library
Dim ULStat%
   ULStat% = cbDeclareRevision(CURRENTREVNUM)
   
   ' Initiate error handling
   '  activating error handling will trap errors like
   '  bad channel numbers and non-configured conditions.
   '  Parameters:
   '    PRINTALL    :all warnings and errors encountered will be printed
   '    DONTSTOP    :if an error is encountered, the program will not stop,
   '                 errors must be handled locally
   
   
   ULStat% = cbErrHandling(PRINTALL, DONTSTOP)
   If ULStat% <> 0 Then Stop
   
   ' If cbErrHandling% is set for STOPALL or STOPFATAL during the program
   ' design stage, Visual Basic will be unloaded when an error is encountered.
   ' We suggest trapping errors locally until the program is ready for compiling
   ' to avoid losing unsaved data during program design.  This can be done by
   ' setting cbErrHandling options as above and checking the value of ULStat%
   ' after a call to the library. If it is not equal to 0, an error has occurred.
   
   ' configure FIRSTPORTA for digital output
   '  Parameters:
   '    BoardDial    :the number used by CB.CFG to describe this board
   '    PortNum%    :the output port
   '    Direction%  :sets the port for input or output
      
   '-------------------------------------------------------------
   ' TAPE TEST BOARD CONFIGURATION
   '-------------------------------------------------------------
    ULStat% = cbDConfigPort(BOARD_IO, FIRSTPORTA, DIGITALIN)
    If ULStat% <> 0 Then Stop
    
    ULStat% = cbDConfigPort(BOARD_IO, FIRSTPORTB, DIGITALIN)
    If ULStat% <> 0 Then Stop
    
    ULStat% = cbDConfigPort(BOARD_IO, FIRSTPORTCL, DIGITALOUT)
    If ULStat% <> 0 Then Stop
      
    ULStat% = cbDConfigPort(BOARD_IO, FIRSTPORTCH, DIGITALOUT)
    If ULStat% <> 0 Then Stop
        
    OutputPort Init_IO, 0
        
'    OutputPort OUTPUT_1A, 1
                 
End Sub
'
'       0:1 ON:OFF
'
Function InputPort(iInp As Integer) As Integer

Dim iBitValue As Integer

Dim ULStat%

Select Case iInp
Case 0 To 7
    ULStat% = cbDIn(BOARD_IO, FIRSTPORTA, iBitValue)
    If ULStat% <> 0 Then Stop
Case 8 To 15
    ULStat% = cbDIn(BOARD_IO, FIRSTPORTB, iBitValue)
    If ULStat% <> 0 Then Stop
Case 16 To 19
    ULStat% = cbDIn(BOARD_IO, FIRSTPORTCL, iBitValue)
    If ULStat% <> 0 Then Stop
Case 20 To 23
    ULStat% = cbDIn(BOARD_IO, FIRSTPORTCH, iBitValue)
    If ULStat% <> 0 Then Stop
End Select

Dim iBit As Integer

Select Case iInp
Case INPUT_1A
                    iBit = 0
Case INPUT_2A
                    iBit = 1
Case INPUT_3A
                    iBit = 2
Case INPUT_4A
                    iBit = 3
Case INPUT_5A
                    iBit = 4
Case INPUT_6A
                    iBit = 5
Case INPUT_7A
                    iBit = 6
Case INPUT_8A
                    iBit = 7
Case INPUT_1B
                    iBit = 0
Case INPUT_2B
                    iBit = 1
Case INPUT_3B
                    iBit = 2
Case INPUT_4B
                    iBit = 3
Case INPUT_5B
                    iBit = 4
Case INPUT_6B
                    iBit = 5
Case INPUT_7B
                    iBit = 6
Case INPUT_8B
                    iBit = 7

Case INPUT_1C
                    iBit = 0
Case INPUT_2C
                    iBit = 1
Case INPUT_3C
                    iBit = 2
Case INPUT_4C
                    iBit = 3
Case INPUT_5C
                    iBit = 0
Case INPUT_6C
                    iBit = 1
Case INPUT_7C
                    iBit = 2
Case INPUT_8C
                    iBit = 3

Case Else

End Select
        
InputPort = (iBitValue And 2 ^ iBit) / 2 ^ iBit

DoEvents

End Function
'
'
'
Sub OutputPort(iOut As Integer, iState As Integer)

Dim iPort As Integer, iBit As Integer

Static iOutWordA As Integer
Static iOutWordB As Integer
Static iOutWordCL As Integer
Static iOutWordCH As Integer
   
Dim i As Integer
  
OUTPUT_Status(iOut) = iState

Dim ULStat%
Select Case iOut
Case Init_IO
                    'iOutWordA = 255
                    'ULStat% = cbDOut(BOARD_IO, FIRSTPORTA, iOutWordA)
                    'If ULStat% <> 0 Then Stop
                    
                    iOutWordCL = 255
                    ULStat% = cbDOut(BOARD_IO, FIRSTPORTCL, iOutWordCL)
                    If ULStat% <> 0 Then Stop
                    
                    iOutWordCH = 255
                    ULStat% = cbDOut(BOARD_IO, FIRSTPORTCH, iOutWordCH)
                    If ULStat% <> 0 Then Stop
                                        
                    For i = 0 To 24
                        OUTPUT_Status(i) = 0
                    Next i
                    Exit Sub
Case 55
                    iOutWordA = 255
                    ULStat% = cbDOut(BOARD_IO, FIRSTPORTA, iOutWordA)
                    If ULStat% <> 0 Then Stop
                    
                    iOutWordB = 255
                    ULStat% = cbDOut(BOARD_IO, FIRSTPORTB, iOutWordB)
                    If ULStat% <> 0 Then Stop
                    
                    iOutWordCL = 255
                    ULStat% = cbDOut(BOARD_IO, FIRSTPORTCL, iOutWordCL)
                    If ULStat% <> 0 Then Stop
                    
                    iOutWordCH = 255
                    ULStat% = cbDOut(BOARD_IO, FIRSTPORTCH, iOutWordCH)
                    If ULStat% <> 0 Then Stop
                    
                  
                    For i = 0 To 24
                        OUTPUT_Status(i) = 0
                    Next i
                    Exit Sub
                    
Case OUTPUT_1A
                    iBit = 1: iPort = FIRSTPORTA
Case OUTPUT_2A
                    iBit = 2: iPort = FIRSTPORTA
Case OUTPUT_3A
                    iBit = 3: iPort = FIRSTPORTA
Case OUTPUT_4A
                    iBit = 4: iPort = FIRSTPORTA
Case OUTPUT_5A
                    iBit = 5: iPort = FIRSTPORTA
Case OUTPUT_6A
                    iBit = 6: iPort = FIRSTPORTA
Case OUTPUT_7A
                    iBit = 7: iPort = FIRSTPORTA
Case OUTPUT_8A
                    iBit = 8: iPort = FIRSTPORTA
                                        
Case OUTPUT_1B
                    iBit = 1: iPort = FIRSTPORTB
Case OUTPUT_2B
                    iBit = 2: iPort = FIRSTPORTB
Case OUTPUT_3B
                    iBit = 3: iPort = FIRSTPORTB
Case OUTPUT_4B
                    iBit = 4: iPort = FIRSTPORTB
Case OUTPUT_5B
                    iBit = 5: iPort = FIRSTPORTB
Case OUTPUT_6B
                    iBit = 6: iPort = FIRSTPORTB
Case OUTPUT_7B
                    iBit = 7: iPort = FIRSTPORTB
Case OUTPUT_8B
                    iBit = 8: iPort = FIRSTPORTB
                                        
Case OUTPUT_1
                    iBit = 1: iPort = FIRSTPORTCL
Case OUTPUT_2
                    iBit = 2: iPort = FIRSTPORTCL
Case OUTPUT_3
                    iBit = 3: iPort = FIRSTPORTCL
Case OUTPUT_4
                    iBit = 4: iPort = FIRSTPORTCL
Case OUTPUT_5
                    iBit = 1: iPort = FIRSTPORTCH
Case OUTPUT_6
                    iBit = 2: iPort = FIRSTPORTCH
Case OUTPUT_7
                    iBit = 3: iPort = FIRSTPORTCH
Case OUTPUT_8
                    iBit = 4: iPort = FIRSTPORTCH
Case Else
End Select
    
iBit = 2 ^ (iBit - 1)

Dim iMask As Integer
iMask = 255 Xor iBit

Select Case iPort
Case FIRSTPORTA
                Select Case iState
                Case 1
                        iOutWordA = iOutWordA And iMask     '   TURN ON
                Case 0
                        iOutWordA = iOutWordA Or iBit       '   TURN OFF
                End Select
                ULStat% = cbDOut(BOARD_IO, FIRSTPORTA, iOutWordA)
                If ULStat% <> 0 Then Stop
Case FIRSTPORTB
                Select Case iState
                Case 1
                        iOutWordB = iOutWordB And iMask     '   TURN ON
                Case 0
                        iOutWordB = iOutWordB Or iBit       '   TURN OFF
                End Select
                ULStat% = cbDOut(BOARD_IO, FIRSTPORTB, iOutWordB)
                If ULStat% <> 0 Then Stop
Case FIRSTPORTCH
                Select Case iState
                Case 1
                        iOutWordCH = iOutWordCH And iMask     '   TURN ON
                Case 0
                        iOutWordCH = iOutWordCH Or iBit       '   TURN OFF
                End Select
                ULStat% = cbDOut(BOARD_IO, FIRSTPORTCH, iOutWordCH)
                If ULStat% <> 0 Then Stop

Case FIRSTPORTCL
                Select Case iState
                Case 1
                        iOutWordCL = iOutWordCL And iMask     '   TURN ON
                Case 0
                        iOutWordCL = iOutWordCL Or iBit       '   TURN OFF
                End Select
                ULStat% = cbDOut(BOARD_IO, FIRSTPORTCL, iOutWordCL)
                If ULStat% <> 0 Then Stop
                
End Select
    
End Sub
