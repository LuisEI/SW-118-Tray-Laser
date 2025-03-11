Attribute VB_Name = "Motion_Module"


Public csr As Integer          'Communication Status Register
Public moveComplete As Integer 'Move complete status


'Global Modal variables
Public errorCode As Long       'Modal ErrorCode
Public commandID As Integer    'Command ID for modal error handling
Public resourceID As Integer   'Resource ID for modal error handling
Public error As Long

Option Explicit


Public Sub MoveToTarget(axis As Integer, targetPosition As Long)

 On Error GoTo Errorhandler
   
    'Load the operation mode - absolute position
    error = flex_set_op_mode(BOARD_ID, axis, NIMC_ABSOLUTE_POSITION)
 
    'Load a target position of 20000 counts or steps
    error = flex_load_target_pos(BOARD_ID, axis, targetPosition, &HFF)
    
    'Start the motion
    error = flex_start(BOARD_ID, axis, 0)
        
    Do
        'Check the move complete status
        error = flex_check_move_complete_status(BOARD_ID, axis, 0, moveComplete)
        
        'Check the modal errors
        flex_read_csr_rtn BOARD_ID, csr
        If (csr And NIMC_MODAL_ERROR_MSG) Then
            'Stop the Motion
            flex_stop_motion BOARD_ID, axis, NIMC_DECEL_STOP, 0
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            CheckError (errorCode)
        End If
        
        DoEvents
    Loop Until moveComplete
    
    Exit Sub    'Exit the Sub
            
'////////////////////////////////////////////////////////////////////////
' Error Handling
'
Errorhandler:
    'First check for modal errors
    flex_read_csr_rtn BOARD_ID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn BOARD_ID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

Public Sub FindReverseLimit(axis As Integer)

On Error GoTo Errorhandler
      
    Dim position As Long
    
    Dim found As Integer
    Dim finding As Integer
    
    Dim axisStatus As Integer
                                
    Dim inputVector
    inputVector = 32
                
    error = flex_find_reference(BOARD_ID, axis, 0, NIMC_FIND_REVERSE_LIMIT_REFERENCE)
    CheckError (error)

    Do
        error = flex_read_pos_rtn(BOARD_ID, axis, position)
        CheckError (error)
        
        error = flex_check_reference(BOARD_ID, axis, 0, found, finding)
        CheckError (error)
                
        error = flex_read_axis_status_rtn(BOARD_ID, axis, axisStatus)
        CheckError (error)
        
        DoEvents
    Loop Until (finding = 0)

    error = flex_reset_pos(BOARD_ID, axis, 0, 0, inputVector)
    CheckError (error)

    MsgBox "flex_find_reference Reverse Limit " & position, vbInformation, "ATC Tray Laser System"

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn BOARD_ID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn BOARD_ID, csr
            
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

Public Sub Initialize_Controller()

Dim status As Long
Dim settingsName As String

status = flex_initialize_controller(BOARD_ID, settingsName)
If (status <> NIMC_noError) Then
    MsgBox "National Instruments Motion Initialization Error", vbCritical, "ATC Tray Laser System"
End If

End Sub

Public Sub Load_Parameters()

On Error GoTo Errorhandler
   
Dim axis As Integer
   
Set FR_Database = OpenDatabase(ATC_LASER_BD)
        
Dim sSQL As String
For axis = 1 To 3
        
    sSQL = "SELECT * FROM  [TBL AXIS] WHERE [AXIS_ID] = " & axis
    Set FR_Table = FR_Database.OpenRecordset(sSQL)
    If (FR_Table.RecordCount <> 0) Then
    
            error = flex_load_rpm(BOARD_ID, axis, FR_Table.Fields("[Velocity]"), &HFF)
            CheckError (error)

            error = flex_load_rpsps(BOARD_ID, axis, NIMC_ACCELERATION, FR_Table.Fields("[Accel]"), &HFF)
            CheckError (error)
            
            error = flex_load_rpsps(BOARD_ID, axis, NIMC_DECELERATION, FR_Table.Fields("[Decel]"), &HFF)
            CheckError (error)
    End If

    'Check the modal errors
    If csr And NIMC_MODAL_ERROR_MSG Then
        error = flex_stop_motion(BOARD_ID, axis, NIMC_DECEL_STOP, 0) 'Stop the Motion
        flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
        CheckError (errorCode)
    End If

Next axis
    
FR_Table.Close
FR_Database.Close
        
        
    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn BOARD_ID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn BOARD_ID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

Public Sub StopMotion(axis As Integer)

On Error GoTo Errorhandler
            'Stop the Motion
    flex_stop_motion BOARD_ID, axis, NIMC_DECEL_STOP, 0
            
    flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
     
    Exit Sub    'Exit the Sub
            
'////////////////////////////////////////////////////////////////////////
' Error Handling
'
Errorhandler:
    'First check for modal errors
    flex_read_csr_rtn BOARD_ID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn BOARD_ID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If
    
End Sub

Public Sub DisableHome()

On Error GoTo Errorhandler
            
    Dim homemap%
    homemap% = 0
    homemap% = BinaryToDecimal("0")
    
    error = flex_enable_home_inputs(BOARD_ID, homemap%)
       
    CheckError (error)
   
    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn BOARD_ID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn BOARD_ID, csr
            
        Loop Until Not (csr And NIMC_MODAL_ERROR_MSG)
        
    Else        'Display regular error
        nimcDisplayError Err.number, 0, 0
    End If

End Sub

Public Sub FindReverseLimitAxis()

On Error GoTo Errorhandler
    Dim axis As Integer
      
    Dim position As Long
    
    Dim found(3) As Integer
    Dim finding(3) As Integer
    
    Dim axisStatus As Integer
                                
    Dim inputVector
    inputVector = 32
                
    For axis = 1 To 3
        error = flex_find_reference(BOARD_ID, axis, 0, NIMC_FIND_REVERSE_LIMIT_REFERENCE)
        CheckError (error)
        Do
                error = flex_read_pos_rtn(BOARD_ID, axis, position)
                CheckError (error)
                
                error = flex_check_reference(BOARD_ID, axis, 0, found(axis), finding(axis))
                CheckError (error)
                        
                error = flex_read_axis_status_rtn(BOARD_ID, axis, axisStatus)
                CheckError (error)
                
                DoEvents
        Loop Until (finding(1) = 0 And finding(2) = 0 And finding(3) = 0)

    error = flex_reset_pos(BOARD_ID, axis, 0, 0, inputVector)
    CheckError (error)

    Next axis
    
    MsgBox "flex_find_reference Reverse Limit " & position, vbInformation, "ATC Tray Laser System"

    Exit Sub
            
'Error Handling
Errorhandler:

    'First check for modal errors
    flex_read_csr_rtn BOARD_ID, csr
    If csr And NIMC_MODAL_ERROR_MSG Then
        Do
            'Get the command ID, resource and the error code of the modal
            '  error from the error stack on the board
            flex_read_error_msg_rtn BOARD_ID, commandID, resourceID, errorCode
            nimcDisplayError errorCode, commandID, resourceID
            
            'Read the Communication Status Register
            flex_read_csr_rtn BOARD_ID, csr
            
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

